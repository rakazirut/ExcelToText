import xlrd
import os.path

print("This tool will read your test cases from excel and write to a text file which is properly formatted "
      "to be imported automatically to DevOps.\nPlease refer to the formatting guide of test case documents to "
      "ensure this tool works properly.\nThe Output.txt file will be created wherever this script is executed.\n")
fn = input("Enter the .xlsx filename (don't include .xlsx): ")  # user can enter filename to read from
# data validation for file not found
if os.path.exists(os.getcwd() +'\\' + fn +'.xlsx'):  # file was found
    wb = xlrd.open_workbook(os.path.join(os.getcwd(), fn + '.xlsx'))  # opens the filename provided in cwd
    wb.sheet_names()  # gets sheet names (probably don't need)
    sh = wb.sheet_by_index(0)  # looking at first sheet
    i = 0  # first row
    my_file = open("Output.txt", "w")  # opens and readies text file

    while sh.cell(i, 0).value != 0:  # put a 0 at the end of excel document in first column to end without error
        test_d = sh.cell(i, 0).value  # load value from cell (i,0) i row first column
        step_d = sh.row_values(i, 1, 2)  # load values i row second column
        result_d = sh.row_values(i, 2, 3)  # load values i row third column
        if test_d == '':
            DB1 = " ".join(step_d)  # build test step portion
            DB2 = " ".join(result_d)  # build test result portion
            my_file.write('|'+DB1 + '|' + DB2 + "\n")  # write data to file
            if step_d == '':
                print("")
        else:
            DB1 = "\n[" + test_d + "]\n"
            my_file.write(DB1)
        i += 1  # increment i for next row
        # print("Line " + str(i) + " was written.")  # console log of progress
    my_file.close()  # release resources
    print("All done. " + str(i) + " lines written.")
else:  # file not found
    print("File not found.")
