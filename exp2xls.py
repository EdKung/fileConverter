import os
import sys
import xlwt
import xlrd

def createFolder(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def run(src_folder, output_folder):
    FORMAT_EXP = '.exp'
    FORMAT_XLS = '.xls'
    count = 1

    createFolder(output_folder)

    if not os.path.isdir(src_folder):
        print('source directory is not found')
        return

    for name in os.listdir(src_folder):
        fileName = os.path.join(src_folder, name)
        if os.path.isfile(fileName) and not name.startswith('.'):
            print('parsing File #', count)
            count = count + 1

            book = xlwt.Workbook()
            ws = book.add_sheet('First Sheet')  # Add a sheet
            f = open(fileName, 'r')
            data = f.readlines()                # read all lines at once

            for i in range(len(data)):
                row = data[i].split('\t')
                for j in range(len(row)):
                    ws.write(i, j, row[j])

            newname = (name.split(FORMAT_EXP)[0])
            book.save(output_folder + '/' + newname + FORMAT_XLS)
            f.close()
    print('parse Finished')
