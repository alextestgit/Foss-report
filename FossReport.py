# fossvenv/Lib/site-packages/xls2xlsx/htmlxls2xlsx.py update line 37
import csv

import fossOutlook
import os
import glob
# import subprocess
import extract_msg
import datetime
import win32com.client

import xlwings as xw

# from xls2xlsx import XLS2XLSX

INPUT_FOLDER = "\\sources\\"
OUTPUT_FOLDER = "output"
SECTIONS = {"Critical vulnerabilities": "Critical", "High vulnerabilities": "High"}
# critical = []
# high = []
new_report = []
xlsx_files = []
csv_files = []
csv_reader = None


def remove_changes(value):
    for char in "[]'":
        value = value.replace(char, '')
    return value


def openExcel(full_filename):
    global ws
    ws = xw.Book(full_filename).sheets[0]


def get_section_row(start_row, name):
    for row in range(start_row, 1000):
        cell = f'A{row}'
        value = str(ws.range(cell).value)
        if name in value:
            print(f"---------- find {name} in row {row} ---------")
            break
    return row + 3


def pint_data(row, name):
    global critical, high
    for row_ in range(row, 100):
        cell = f'A{row_}:F{row_}'
        row_value = str(ws.range(cell).value)
        row_value = row_value.split(",")
        row0 = remove_changes(row_value[0])
        if "None" not in row0:
            print(row0)
            if "Critical" in name:
                critical.append(row_value)
            else:
                high.append(row_value)
        else:
            break


def work_with_csv(filename):
    global csv_reader
    with open(filename, 'r') as csv_file:
        csv_reader = csv.reader(csv_file, dialect='excel')
        for section_name, severity in SECTIONS.items():
            move2section_date(section_name)
            product_name = filename.split("\\")[-1].split("_")[0]
            get_section_data(product_name,severity)
            "".split()


def move2section_date(section_name):
    for row in csv_reader:
        if section_name in row[0]:
            next(csv_reader)
            next(csv_reader)
            return


def get_section_data(product, severity):
    for row in csv_reader:
        if row[0] == '':
            return
        report_row = row
        report_row.insert(0, product)
        report_row.insert(len(row), severity)
        new_report.append(report_row)

    # def work_with_csv(filename):
    open_csv_file(filename)
    for section_name in SECTIONS:
        move2section_date(section_name)
        get_section_data()


def work_with_excel(file):
    openExcel(file)
    section_row = 1
    for section_name in SECTIONS:
        section_row = get_section_row(section_row, section_name)
        pint_data(section_row, section_name)


#
# def convertxls2xlsx(in_filename, out_filename):
#     x2x = XLS2XLSX(in_filename)
#     x2x.to_xlsx(out_filename)

# def open_message(file_):
#     # outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#     # file_path = SOURCE_FOLDER + file_
#     # msg = outlook.OpenSharedItem(file_path)
#     msg = extract_msg.Message(SOURCE_FOLDER + file_)
#     # https://stackoverflow.com/questions/26322255/parsing-outlook-msg-files-with-python
#     return msg


def extract_excels_from_msg():
    for file_ in os.listdir(SOURCE_FOLDER):
        if file_.endswith(".msg"):
            msg = extract_msg.Message(SOURCE_FOLDER + file_)
            attachments = msg.attachments
            for att in attachments:
                if att.extension == ".xls":
                    att_name = xls_name(msg)
                    xlsx_files.append(att_name)
                    with open(att_name, 'wb') as fl:
                        fl.write(att.data)


def xls_name(msg):
    date = '_'.join(msg.date.split()[:-2]).replace(',', '')
    out_name = msg.subject.split("-")[1].strip().replace(':', '')
    return f'{SOURCE_FOLDER}{out_name}_{date}.xls'


def is_folder_empty(path: str) -> bool:
    return len(os.listdir(path)) == 0


# def convert_xls2xlsx():
#     # https://stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-line
#     global xlsx_files
#     xls_files = glob.glob(SOURCE_FOLDER + "*.xls")
#     xls_files = [xls_file for xls_file in xls_files if "~" not in xls_file]
#     for xls_file in xls_files:
#         xlsx_file = os.path.splitext(xls_file)[0] + '.xlsx'
#         cmd = f'{SOURCE_FOLDER}XlsToCsv1.vbs {xls_file} {xlsx_file}'
#         returned_output = os.system(cmd)
#         print(f'Converter output: {returned_output}')
#         xlsx_files.append(xls_file)

def convert_xls2csv():
    cmd = f'{directory_path}\XlsToCsv.vbs {SOURCE_FOLDER}'
    returned_output = os.system(cmd)
    print(returned_output)


def get_csv_files() -> list:
    for file in os.listdir(SOURCE_FOLDER):
        if file.endswith(".csv"):
            csv_files.append(f'{SOURCE_FOLDER}\\{file}')
    # return csv_files


# --------- MAIN ------------
directory_path = os.getcwd()
SOURCE_FOLDER = directory_path + INPUT_FOLDER

extract_excels_from_msg()
convert_xls2csv()
get_csv_files()
# for file in xlsx_files:

for file in csv_files:
    work_with_csv(file)

print("\n\n\n--- data ")
for row in new_report:
    print(row)
