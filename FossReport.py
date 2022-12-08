# fossvenv/Lib/site-packages/xls2xlsx/htmlxls2xlsx.py update line 37
import os
import glob
import subprocess

import xlwings as xw

# from xls2xlsx import XLS2XLSX

INPUT_FOLDER = "\\sources\\"
OUTPUT_FOLDER = "output"
SECTIONS = ["Critical vulnerabilities", "High"]
critical = []
high = []
xlsx_files =[]

def remove_changes(value):
    for char in "[]'":
        value = value.replace(char, '')
    return value


#
# def convertxls2xlsx(in_filename, out_filename):
#     x2x = XLS2XLSX(in_filename)
#     x2x.to_xlsx(out_filename)


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


def convert_xls2xlsx():
    # https://stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-line
    global xlsx_files
    xls_files = glob.glob(SOURCE_FOLDER + "*.xls")
    xls_files = [xls_file for xls_file in xls_files if "~" not in xls_file]
    for xls_file in xls_files:
        xlsx_file = os.path.splitext(xls_file)[0] + '.xlsx'
        # XlsToCsv1.vbs AIF_9.22_1.xls   AIF_test.xlsx
        cmd = f'{SOURCE_FOLDER}XlsToCsv1.vbs {xls_file} {xlsx_file}'
        # returned_output = subprocess.check_output(cmd)
        returned_output = os.system(cmd)
        # print(f'Converter output: {returned_output.decode("utf-8")}')
        print(f'Converter output: {returned_output}')
        xlsx_files.append(xls_file)


# --------- MAIN ------------
directory_path = os.getcwd()
SOURCE_FOLDER = directory_path + INPUT_FOLDER
# in_file, convert_file = "AIF_9.22_1.xls", "AIF_9.22_1.xlsx"
# in_file, convert_file = "WSF_9.22.xls", "WSF_9.22.xlsx"
# in_file = FOLDER + in_file
# convert_file = FOLDER + convert_file

# convertxls2xlsx(in_file, convert_file)
convert_xls2xlsx()
for file in xlsx_files:
    print(f'File name {file}')
    work_with_excel(file)

print("--- CRITICAL ")
for row in critical:
    print(row)
print("---------- HIGH ")
for row in high:
    print(row)
