import xlrd
import os.path

wb = xlrd.open_workbook(os.path.join(os.getcwd(), 'Demo.xlsx'))
wb.sheet_names()
sh = wb.sheet_by_index(0)
i = 0
my_file = open("Output.txt", "w")

while sh.cell(i,0).value != 0:
    Load = sh.cell(i,0).value
    step_d = sh.row_values(i, 1, 2)
    result_d = sh.row_values(i, 2, 3)
    DB1 = Load+" "+(" ".join(step_d))
    DB2 = " ".join(result_d)
    my_file.write(DB1 + ' | ' + DB2 + "\n")
    i += 1
my_file.close()