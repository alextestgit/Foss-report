# fossvenv/Lib/site-packages/xls2xlsx/htmlxls2xlsx.py update line 37
import csv

import fossOutlook
import os
import extract_msg
import xlwings as xw

INPUT_FOLDER = "\\sources\\"
OUTPUT_FOLDER = "output"
SECTIONS = {"Critical vulnerabilities": "Critical", "High vulnerabilities": "High"}
column_names = ["Product", "FOSS name", "FOSS version", "Latest clean version", "Nearest clean version", "Defect #",
                "Comments", "Communication Platform", "Severity"]

column_size = [11, 30, 13, 27, 35, 11, 30, 25, 10]

report_name = "Summary_Foss_Report_"

new_report, xlsx_files, csv_files = [], [], []
csv_reader, msg_date = None, None


def create_report():
    wb = xw.Book()
    ws = wb.sheets[0]
    row_num = 1
    # Write the column names to the Excel file
    for i, column_name in enumerate(column_names):
        ws.range((row_num, i + 1)).value = column_name

    # Set the font of the column names to bold
    ws.range((row_num, 1), (row_num, len(column_names))).api.Font.Bold = True
    row_num += 2
    # Write the dates to the Excel file
    for i, date in enumerate(new_report, row_num):
        ws.range((i, 1)).value = date

    # Set the width of each column to the maximum length of the data in that column
    for i, column_width in enumerate(column_size):
        ws.range((row_num, i + 1)).column_width = column_width

    ws.used_range.api.AutoFilter(Field := 1)

    report_name_final = f'{SOURCE_FOLDER}\\{report_name}_{msg_date}.xlsx'
    wb.save(report_name_final)


def work_with_csv(filename):
    global csv_reader
    with open(filename, 'r') as csv_file:
        csv_reader = csv.reader(csv_file, dialect='excel')
        for section_name, severity in SECTIONS.items():
            move2section_date(section_name)
            product_name = filename.split("\\")[-1].split("_")[0]
            get_section_data(product_name, severity)
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
    global msg_date
    msg_date = '_'.join(msg.date.split()[:-2]).replace(',', '')
    out_name = msg.subject.split("-")[1].strip().replace(':', '')
    return f'{SOURCE_FOLDER}{out_name}_{msg_date}.xls'


def is_folder_empty(path: str) -> bool:
    return len(os.listdir(path)) == 0


def convert_xls2csv():
    cmd = f'{directory_path}\XlsToCsv.vbs {SOURCE_FOLDER}'
    returned_output = os.system(cmd)
    print(returned_output)


def get_csv_files() -> list:
    for file in os.listdir(SOURCE_FOLDER):
        if file.endswith(".csv"):
            csv_files.append(f'{SOURCE_FOLDER}\\{file}')
    # return csv_files


def createNewReport():
    pass


# --------- MAIN ------------
directory_path = os.getcwd()
SOURCE_FOLDER = directory_path + INPUT_FOLDER

extract_excels_from_msg()
convert_xls2csv()
get_csv_files()
for file in csv_files:
    work_with_csv(file)
create_report()

# print("\n\n\n--- data ")
# for row in new_report:
#     print(row)
