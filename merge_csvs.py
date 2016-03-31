import os
import xlrd
import xlsxwriter
import argparse

# http://www.tutorialspoint.com/python/os_walk.htm
# recursively get all files inside any subdirectories given a directory
def walk_files(root_dir):
    file_names = list()
    for root, dirs, files in os.walk(root_dir, topdown=True):
        for name in files:
            file_names.append(os.path.join(root, name))
    return file_names

# https://github.com/python-excel/xlrd
def read_xls(xls_file, sheet_idx):
    book = xlrd.open_workbook(xls_file)
    return book.sheet_by_index(sheet_idx)

# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/
def read_sheet(xls_sheet):
    rows = list()
    for row_idx in range(0, xls_sheet.nrows):
        cells = list()
        for col_idx in range(0, xls_sheet.ncols):
            cell_obj = xls_sheet.cell(row_idx, col_idx)
            cells.append(cell_obj.value)
        rows.append(cells)
    return rows

# remove the first 3 rows and the last row
def clean_rows(rows):
    return rows[3:-1]

# all files have same header, on third row
def get_header(rows):
    return rows[2:3]

# http://xlsxwriter.readthedocs.org/example_demo.html
# http://stackoverflow.com/questions/23813237/xlrd-xlwt-in-python-how-to-copy-an-entire-row
def write_xls(xls_file, sheet_name, data_rows):
    workbook = xlsxwriter.Workbook(xls_file)
    worksheet = workbook.add_worksheet(sheet_name)
    for row_idx, row in enumerate(data_rows):
        for col_idx, col in enumerate(row):
            worksheet.write(row_idx, col_idx, col)
    workbook.close()

def parse_arguments():
    parser = argparse.ArgumentParser(description='Merge Excel data.')
    parser.add_argument('directory', help='The root directory to recursively search for Excel files.')
    parser.add_argument('filename', help='The name of the merged Excel file to output.')
    return parser.parse_args()

def main():
    # args = parse_arguments()
    #root_dir = args.directory
    root_dir = "S:\Projects\Open\KCFCCC - KConnect Data, Research, & Evaluation\Data Files\COMPLETE PUBLIC Building Files"
    #workbook_name = args.filename
    workbook_name = "PUBLIC.xlsx"
    sheet_name = 'Merged Data'

    # recursively search for files from directory
    file_names = walk_files(root_dir)
    file_names = file_names[0:3]
    target_rows = list()

    # get header of merged file
    header = get_header(read_sheet(read_xls(file_names[0], 0)))[0]
    target_rows.append(header)

    # get rows of data
    for xls_file in file_names:
        sheet = read_xls(xls_file, 0)
        rows = read_sheet(sheet)
        cleaned_rows = clean_rows(rows)
        cleaned_rows = cleaned_rows[0:2]
        target_rows = target_rows + cleaned_rows

    # output the file
    write_xls(workbook_name, sheet_name, target_rows)

if __name__ == "__main__":
    main()
