from glob import glob
import re
import os
import csv
from time import time
from datetime import datetime
import win32com.client as win32
from zipfile import ZipFile
from win32com.client import constants

# Create list of paths to .doc files
paths = glob('C:\\DocRename\\Docx\\*.docx', recursive=True)
paths_zip = glob('C:\\DocRename\\*.zip', recursive=True)

FILE_FORMAT_PDF_WORD = 17


def save_as_docx(path, old_name, new_name, pdf_name):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    # new_file_abs = os.path.abspath(path)
    # new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    print('Converting ' + old_name + ' to ' + new_name)
    word.ActiveDocument.SaveAs('C:\\DocRename\\Docx\\' + new_name, FileFormat=constants.wdFormatXMLDocument)

    print('Converting ' + new_name + ' to ' + pdf_name)
    word.ActiveDocument.SaveAs('C:\\DocRename\\Docx\\' + pdf_name, FileFormat=FILE_FORMAT_PDF_WORD)

    doc.Close(False)


def convert_to_pdf(file_in):
    # Use the correct COM server
    print('Converting DocX to PDF: ' + file_in)

    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(file_in)
    doc.Activate()

    # Rename path with .pdf
    new_file_abs = os.path.abspath(file_in)
    new_file_abs = re.sub(r'\.\w+$', '.pdf', new_file_abs)

    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=FILE_FORMAT_PDF_WORD)
    doc.Close(False)


def unzip(path):
    with ZipFile(path, 'r') as zipObject:
        listOfFileNames = zipObject.namelist()
        for fileName in listOfFileNames:
            # print(fileName)
            if fileName.endswith('.doc'):
                # Extract a single file from zip
                print('Extracting ' + fileName)
                zipObject.extract(fileName, 'c:\\DocRename\\Doc\\')
                # print('All the python files are extracted')


def main():
    print('starting')
    config = config_file()

    log = prepare_log()

    for zip_name in config["Rename"]:
        zip_name_path = 'C:\\DocRename\\Zip\\' + zip_name[2] + '.zip'
        # print('unzipping ' + zip_name_path)
        unzip(zip_name_path)
        doc_name_path = 'C:\\DocRename\\Doc\\' + zip_name[0]
        # print(doc_name_path)
        save_as_docx(doc_name_path, zip_name[0], zip_name[1], zip_name[3])
        # docx_name_path = 'C:\\DocRename\\Docx\\' + zip_name[0]
        # convert_to_pdf(docx_name_path)
        log.write(f"{zip_name[2]}.zip;{zip_name[0]};{zip_name[1]};{zip_name[3]};Unzipped;Renamed\n")

    log.close()

    '''
    for path in paths:
    '''

    print('finished')


def config_file():

    cfg_dict = {}
    cfg_dict["Rename"] = []

    csv_filename = 'c:\\DocRename\\DocRename.csv'
    with open(csv_filename, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # print(row["Name"])
            cfg_dict["Rename"].append([row["FileName"], row["FileNameNew"], row["JobID"], row["PDFName"]])

    # print(cfg_dict)
    return cfg_dict


def prepare_log():

    log_time = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    log_name = f"Zip_DocX_Validation_{log_time}.csv"
    log_path = os.path.join("C:\\DocRename\\Log\\", log_name)

    log = open(log_path, "w+", encoding="utf-8")

    log.write("sep=;\n")
    log.write("ZipName;DocName;DocXName;PDFName;ZipStatus;RenameStatus\n")

    print(f"Logging in file {log_name}")
    return log


if __name__ == "__main__":
    main()

