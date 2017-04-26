import os
import sys
import xlrd
import openpyxl

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
    FORMAT_XLSX = '.xlsx'
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

            wb = openpyxl.Workbook()
            ws = wb.active
            with open(fileName, 'r+') as f:
                data = f.readlines()
                for i in range(len(data)):
                    row = data[i].split('\t')
                    line = []
                    for j in range(len(row)):
                        if isfloat(row[j]):
                            line.append(float(row[j]))
                        else:
                            line.append(row[j])
                    ws.append(line)

            newname = (name.split(FORMAT_EXP)[0])
            dst_filename = output_folder + '/' + newname + FORMAT_XLSX
            wb.save(filename = dst_filename)
            f.close()
    print('parse Finished')
