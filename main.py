def create_template(path):
    print('Hi')
    today = datetime.now()
    today = today.strftime('%y%y%m%d%H%M%S')
    print(today)
    temp_path = os.path.join(path, today)
    #temp_path = today
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(f'{temp_path}.xlsx')
    worksheet0 = workbook.add_worksheet('ATR')  # Defaults to Sheet1.
    worksheet1 = workbook.add_worksheet('ESS')  # Data.
    worksheet2 = workbook.add_worksheet('Statistics')  # Defaults to Sheet

    # Some data we want to write to the worksheet.
    Tests_List = ['Temp', 'SN', 'Output Power @ P1dBCP', 'Output Power Control Range/Resolution, FWD PWR Ind',
                  'Output IP3', 'LO Carrier Leakage', 'Sideband Suppression',
                  'Frequency Accuracy and Stability', 'A1 - Noise Figure vs. Gain', 'A1 - Gain variability',
                  'A1 - Image Suppression vs. Gain', 'Spurious',
                  'A2 - Noise Figure vs. Gain', 'A2 - Gain variability', 'A2 - Image Suppression vs. Gain',
                  'Average Power Consumption', 'Input Voltage', 'Digital Tests'
                  ]

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    # col = 0

    # Iterate over the data and write it out row by row.
    for index in range(3) :
        for i in range(len(Tests_List)) :
            worksheet0.write(row, i, Tests_List[i])
            worksheet1.write(row, i, Tests_List[i])
            worksheet2.write(row, i, Tests_List[i])
            # col += 1

    workbook.close()

    return today, temp_path


def excel_fun_read(file_name, template_name, template_location):
    for list_number in range(1, 6):
        inputWorkbook = xlrd.open_workbook(file_name)
        inputWorksheet = inputWorkbook.sheet_by_index(list_number)
        print(inputWorksheet.nrows, '\n') # <- get rows number starts from 0
        print(inputWorksheet.ncols, '\n') # <- get coloms number starts from 0
        if list_number == 1:
            sn = int(inputWorksheet.cell_value(0, 8)) # <- get SN number
            print(sn)   # <- Indicates which file is open
            Tests_List = ['Temp', 'SN', 'Output Power @ P1dBCP', 'A1 - Output Power Control Range/Resolution, FWD PWR Ind',
                          'A2 - Output Power Control Range/Resolution, FWD PWR Ind', 'Output IP3', 'LO Carrier Leakage', 'Sideband Suppression',
                          'Frequency Accuracy and Stability', 'A1 - Noise Figure vs. Gain', 'A1 - Gain variability',
                          'A1 - Image Suppression vs. Gain', 'Spurious',
                          'A2 - Noise Figure vs. Gain', 'A2 - Gain variability', 'A2 - Image Suppression vs. Gain',
                          'Average Power Consumption', 'Input Voltage', 'Digital Tests'
                          ]
            Tests_locations_B = [] #<- Creation of list
            B_col_expended = [] #<- length of cells
            G_col = []  #<- Creation of list
            B_col = []  #<- where B_col_expended starts
            G_ATP = []
            ESS_1_Cold = []
            ESS_2_Cold = []
            ESS_1_Hot = []
            ESS_2_Hot = []

            for i in range(inputWorksheet.nrows):

            #Follow the H colom check if there is 'PASS'/'FAIL' or empty cell
            #If empty cell skip it until the end of the excel file

                if inputWorksheet.cell_value(i, 7) == 'PASS' or inputWorksheet.cell_value(i, 7) == 'FAIL':
                    Tests_locations_B.append(i)
                    G_col.append(str(inputWorksheet.cell_value(i, 6))) # <- Only in ATR
                    #B_col.append(str(inputWorksheet.cell_value(i - 1, 1)))

            B_col_expended, B_col = Create_2_lists_of_locations(Tests_locations_B, B_col_expended, B_col, list_number)
            sorted_column = sort_list_of_pass_and_fail(B_col_expended, B_col, file_name, list_number)

            print(f'''It's the end of {list_number} loop''')
            write_to_excel(file_name, template_name, template_location, list_number, sorted_column)
    print('''it's the end of the loop''')

    #print(files.head())
    #files.head()

def Create_2_lists_of_locations(Tests_locations_B, B_col_expended, B_col, list_number):
    counter = 1
    for i_sub in range(len(Tests_locations_B)):

        try:
            if list_number == 1:
                if Tests_locations_B[i_sub] == 6 or Tests_locations_B[i_sub] == 110 or Tests_locations_B[i_sub] == 139 or Tests_locations_B[i_sub] == 172 or Tests_locations_B[i_sub] == 224:
                    B_col.append(Tests_locations_B[i_sub])
            if Tests_locations_B[i_sub] + 1 == Tests_locations_B[i_sub + 1]:
                if counter == 1:
                    B_col.append(Tests_locations_B[i_sub])
                counter = counter + 1

            else:
                B_col_expended.append(counter)
                counter = 1
        except:
            print("An exception occurred")

    return B_col_expended, B_col


def sort_list_of_pass_and_fail(B_col_expended, B_col, file_name, list_number):
    Results_col = ['PASS']
    inputWorkbook = xlrd.open_workbook(file_name)
    inputWorksheet = inputWorkbook.sheet_by_index(list_number)

    for index in range(len(B_col)):
        for i in range(B_col[index], B_col[index] + B_col_expended[index]):
            print(f'indicated value: {str(inputWorksheet.cell_value(i, 7))}, index {i}')
            if str(inputWorksheet.cell_value(i, 7)) == 'FAIL' or str(inputWorksheet.cell_value(i, 7)) == 'N/T':
                Results_col[index] = str(inputWorksheet.cell_value(i, 7))

        if len(Results_col) < len(B_col):
            Results_col.append('PASS')


    print('Stop')
    return Results_col


def write_to_excel(file_name, template_name, template_location, list_number, sorted_column):
    print(template_location)

    a = os.path.join(template_location + '.xlsx')
    wb = openpyxl.load_workbook(a)
    print(wb.sheetnames)
    inputWorkbook = xlrd.open_workbook(file_name)

    if list_number == 1:
        inputWorksheet = inputWorkbook.sheet_by_index(1)
        row_number = print(inputWorksheet.nrows, '\n')  # <- get rows number starts from 0


    else:
        inputWorksheet = inputWorkbook.sheet_by_index(2)
        row_number = print(inputWorksheet.nrows, '\n')  # <- get rows number starts from 0
        pass
    wb.save(a, as_template=False)
def main(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
    path = '/home/pi/Desktop/Python/fwdreportsnb'
    template_path = '/home/pi/Desktop/Pycharm/Tempalte'
    template_path = 'D:\Rasberry Pie\Python\Excel\Result'
    #path = input('Input Path location: \n')
    path = 'D:\Rasberry Pie\Python\Excel\Excel'
    excel_files = [f for f in os.listdir(path) if f.endswith('.xlsx')]
    excel_files = sorted(excel_files)
    print(excel_files, '\n')

    time_date, Template_path = create_template(template_path)
    print(time_date)
    for i in range(len(excel_files)):
        full_path = os.path.join(path, excel_files[i])
        print('\n', excel_files[i])
        excel_fun_read(full_path, time_date, Template_path)
        #excel_fun_write(template_path, )



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    import os
    #import pandas as pd
    import xlrd
    import xlsxwriter
    import xlwt
    import xlutils
    import openpyxl
    from xlutils.copy import copy
    from xlwt import Workbook
    from datetime import datetime
    main('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/