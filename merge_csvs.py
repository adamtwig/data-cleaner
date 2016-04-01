import os
import xlrd
import csv
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

# https://docs.python.org/dev/library/csv.html#csv.writer
def write_csv(csv_file, data_rows):
    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(data_rows)

# merge excel files into one
def merge_files(root_dir, workbook_name):
    # recursively search for files from directory
    file_names = walk_files(root_dir)
    target_rows = list()
    # get header of merged file
    header = get_header(read_sheet(read_xls(file_names[0], 0)))[0]
    target_rows.append(header)
    # get rows of data
    for xls_file in file_names:
        sheet = read_xls(xls_file, 0)
        rows = read_sheet(sheet)
        cleaned_rows = clean_rows(rows)
        # concatenate lists
        target_rows = target_rows + cleaned_rows
    # output the file
    write_csv(workbook_name, target_rows)

def parse_arguments():
    parser = argparse.ArgumentParser(description='Merge Excel data.')
    parser.add_argument('directory', help='The root directory to recursively search for Excel files.')
    parser.add_argument('filename', help='The name of the merged Excel file to output.')
    return parser.parse_args()

def main():
    # args = parse_arguments()
    # root_dir = args.directory
    root_dir = "S:\Projects\Open\KCFCCC - KConnect Data, Research, & Evaluation\Data Files"
    root_dir = root_dir + "\COMPLETE PUBLIC Building Files"
    # workbook_name = args.filename
    workbook_name = "PUBLIC.csv"
    merge_files(root_dir, workbook_name)



if __name__ == "__main__":
    main()
