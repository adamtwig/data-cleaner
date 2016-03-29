import os

# http://www.tutorialspoint.com/python/os_walk.htm
def walk_files(root_dir):
    file_names = list()
    for root, dirs, files in os.walk(root_dir, topdown=True):
        for name in files:
            file_names.append(os.path.join(root, name))
    return file_names

# http://www.sitepoint.com/using-python-parse-spreadsheet-data/
# https://pypi.python.org/pypi/xlrd
def read_xls(xls_file):
    data = list()
    with open(xls_file, 'rb') as data_file:
        for line in data_file:
            data.append(line)
    return data

def main():
    root_dir = "S:\Projects\Open\KCFCCC - KConnect Data, Research, & Evaluation\Data Files\COMPLETE PUBLIC Building Files"
    file_names = walk_files(root_dir)
    print(file_names[0])
    # files are xls files

    #for xls_file in file_names:
    #    data = read_xls(xls_file)
    #    print(data[1])

    # ignore first 3 lines, and last line of each xls file_names
    # append all other lines to one giant file



if __name__ == "__main__":
    main()
