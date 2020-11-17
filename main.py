def excel_fun(file_name):
    #excel_file = file_name
    #files = pd.read_excel(file_name)
    #files = pd.read_excel(file_name, "ATR Results")
    inputWorkbook = xlrd.open_workbook(file_name)
    inputWorksheet = inputWorkbook.sheet_by_index(1)
    print(inputWorksheet.nrows, '\n') # <- get rows number starts from 0
    print(inputWorksheet.ncols, '\n') # <- get coloms number starts from 0

    print(int(inputWorksheet.cell_value(0, 8))) # <- get SN number
    #print(files.head())
    #files.head()


def main(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
    # path = '/usr/share/cups/charmaps'
    #path = input('Input Path location: \n')
    path = 'D:\Rasberry Pie\Python\Excel\Excel'
    excel_files = [f for f in os.listdir(path) if f.endswith('.xlsx')]
    print(excel_files, '\n')
    print(excel_files[0])
    i = 0
    a = range(len(excel_files))
    for i in range(len(excel_files)):
        full_path = os.path.join(path, excel_files[i])
        print('\n', excel_files[i])
        excel_fun(full_path)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    import os
    import pandas as pd
    import xlrd
    main('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
