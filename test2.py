import os
import tempfile
from google.cloud import storage
from img2table.document import PDF
from img2table.ocr import VisionOCR
from img2table.document import Image
import camelot
import openpyxl
import pandas as pd

storage_client = storage.Client(project="pde-4-414413")  # Put your project name here


def create_directories():
    directories = ["/tmp/excel/pdfs", "/tmp/pdfs/", "/tmp/pass/", "/tmp/camelot/"]

    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        print(f"Created directory: {directory}")


def upload_to_gcs(local_file_path, gcs_bucket_name, gcs_blob_name):
    storage_client = storage.Client.from_service_account_json(
        json_credentials_path=download_path1
    )
    bucket = storage_client.get_bucket(gcs_bucket_name)
    blob = bucket.blob(gcs_blob_name)
    blob.upload_from_filename(local_file_path)
    print(f"Uploaded {local_file_path} to {gcs_bucket_name}/{gcs_blob_name}")


def upload_folder_to_gcs(local_folder, gcs_bucket_name, gcs_folder):
    for root, dirs, files in os.walk(local_folder):
        for file in files:
            local_file_path = os.path.join(root, file)
            gcs_blob_name = os.path.join(gcs_folder, file)
            upload_to_gcs(local_file_path, gcs_bucket_name, gcs_blob_name)


def get_paths():
    global download_path1
    bucket = storage_client.get_bucket("project-8")  # Put your bucket name here
    blobs = bucket.list_blobs(prefix="pdfs/")
    pass_blob = bucket.list_blobs(prefix="pass/")
    for pas in pass_blob:
        if pas.name.endswith("json"):
            download_path1 = os.path.join("/tmp/", f"{pas.name}")
            pas.download_to_filename(download_path1)
    for blob in blobs:
        if blob.content_type == "application/pdf":
            print(blob.name)
            download_path = os.path.join("/tmp/", f"{blob.name}")
            blob.download_to_filename(download_path)
            print(f"Downloaded to: {download_path}")
            ocr = VisionOCR(api_key=download_path1, timeout=15)
            doc = PDF(download_path, detect_rotation=False, pdf_text_extraction=True)

            extracted_tables = doc.extract_tables(
                ocr=ocr, implicit_rows=False, borderless_tables=True, min_confidence=0
            )
            dest_path = os.path.join("/tmp/excel", f"{blob.name}.xlsx")
            doc.to_xlsx(
                dest=dest_path,
                ocr=ocr,
                implicit_rows=True,
                borderless_tables=True,
                min_confidence=0,
            )

            tables = camelot.read_pdf(download_path, flavor="stream", pages="all")

            excel_file_name = (
                f"{os.path.splitext(os.path.basename(blob.name))[0]}_output.xlsx"
            )
            dest_path = os.path.join("/tmp/camelot", excel_file_name)
            excel_writer = pd.ExcelWriter(dest_path, engine="openpyxl")

            for i, table in enumerate(tables):
                df = table.df
                sheet_name = f"Table_{i+1}"
                df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
            excel_writer.close()


def final(event, context):
    create_directories()
    get_paths()
    upload_folder_to_gcs(
        "/tmp/camelot", "project-8", "camelot"
    )  # Put your bucket name in the middle
    upload_folder_to_gcs(
        "/tmp/excel/pdfs", "project-8", "ocr"
    )  # Put your bucket name middle


final(1, 2)
