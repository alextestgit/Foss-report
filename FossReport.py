# fossvenv/Lib/site-packages/xls2xlsx/htmlxls2xlsx.py update line 37
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
SECTIONS = ["Critical vulnerabilities", "High"]
critical = []
high = []
xlsx_files = []


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


def work_with_excel(file):
    openExcel(file)
    current_row = 1
    for section_name in SECTIONS:
        current_row = get_section_row(current_row, section_name)
        pint_data(current_row, section_name)


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
                    with open(xls_name(msg), 'wb') as fl:
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

def convert_xls2xlsx():
    cmd = f'{directory_path}\XlsToCsv1.vbs {SOURCE_FOLDER}'
    returned_output = os.system(cmd)
    print(returned_output)


def get_csv_files(path: str) -> List[str]:
    csv_files = []
    for file in os.listdir(path):
        if file.endswith(".csv"):
            csv_files.append(file)
    return csv_files

# --------- MAIN ------------
directory_path = os.getcwd()
SOURCE_FOLDER = directory_path + INPUT_FOLDER

extract_excels_from_msg()
convert_xls2xlsx()
for file in xlsx_files:
    print(f'File name {file}')
    work_with_excel(file)

print("\n\n\n--- CRITICAL ")
for row in critical:
    print(row)
print("---------- HIGH ")
for row in high:
    print(row)
