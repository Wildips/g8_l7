import os
import glob
import pytest
from zipfile import ZipFile
from xlrd import open_workbook
from openpyxl import load_workbook
from pypdf import PdfReader

PROJECT_PATH = os.path.dirname(os.path.abspath(__file__))
RESOURCES_PATH = os.path.join(PROJECT_PATH, "resources")
FILES_CONTENT_PART = {
    'xls': {'amount_of_sheets': 0, 'sheets_names_list': [], 'sheet_rows': 0, 'sheet_columns': 0, 'sheet_crossing': ''},
    'xlsx': {'sheet_crossing': ''},
    'txt': {'file_text': ''},
    'pdf': {'amount_of_sheets': 0, 'page_text': ''}}


@pytest.fixture(scope='session', autouse=True)
def create_test_data(tmpdir_factory):
    temp_dir_path = tmpdir_factory.mktemp("data").join("test_archive.zip")
    resources_list = glob.glob(os.path.join(RESOURCES_PATH, "*"))

    with ZipFile(temp_dir_path, "w") as zip_file:
        for res_file in resources_list:
            zip_file.write(os.path.abspath(res_file))

            file_extension = res_file[res_file.find('.'):]

            if file_extension == "xls":
                workbook = open_workbook(res_file)
                FILES_CONTENT_PART['xls']['amount_of_sheets'] = workbook.nsheets
                FILES_CONTENT_PART['xls']['sheets_names_list'] = workbook.sheet_names()
                # only for current example
                sheet = workbook.sheet_by_index(0)
                FILES_CONTENT_PART['xls']['sheet_rows'] = sheet.nrows
                FILES_CONTENT_PART['xls']['sheet_columns'] = sheet.ncols
                FILES_CONTENT_PART['xls']['sheet_crossing'] = sheet.cell_value(9, 2)
            elif file_extension == "xslx":
                workbook = load_workbook(res_file)
                sheet = workbook.active()
                FILES_CONTENT_PART['xslx']['sheet_crossing'] = sheet.cell(row=3, column=2).value
            elif file_extension == "pdf":
                reader = PdfReader(res_file)
                FILES_CONTENT_PART['pdf']['amount_of_sheets'] = len(reader.pages)
                page = reader.pages[1]
                FILES_CONTENT_PART['pdf']['page_text'] = page.extract_text()
            elif file_extension == "txt":
                with open(res_file, 'r') as fn:
                    FILES_CONTENT_PART['pdf']['file_text'] = fn.read()

    return temp_dir_path, FILES_CONTENT_PART
