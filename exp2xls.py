import os
import sys
import six
import xlrd
import xlwt

def createFolder(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

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
            style = xlwt.XFStyle()
            style.num_format_str = 'general'

            with open(fileName, 'r+') as f:
                data = f.readlines()
                for i in range(len(data)):
                    row = data[i].split('\t')
                    for j in range(len(row)):
                        if isfloat(row[j]):
                            ws.write(i, j, float(row[j]), style)
                        else:
                            ws.write(i, j, row[j], style)

            newname = (name.split(FORMAT_EXP)[0])
            book.save(output_folder + '/' + newname + FORMAT_XLS)
            f.close()
    print('parse Finished')
