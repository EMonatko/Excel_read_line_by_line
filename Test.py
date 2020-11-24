import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
Tests_List = ['Temp', 'SN', 'Output Power @ P1dBCP','Output Power Control Range/Resolution, FWD PWR Ind', 'Output IP3', 'LO Carrier Leakage','Sideband Suppression',
                  'Frequency Accuracy and Stability', 'A1 - Noise Figure vs. Gain', 'A1 - Gain variability', 'A1 - Image Suppression vs. Gain', 'Spurious',
                  'A2 - Noise Figure vs. Gain', 'A2 - Gain variability', 'A2 - Image Suppression vs. Gain', 'Average Power Consumption', 'Input Voltage', 'Digital Tests'
                  ]

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

worksheet1 = workbook.add_worksheet('ATR')        # Defaults to Sheet1.
worksheet2 = workbook.add_worksheet('ESS')  # Data.
worksheet3 = workbook.add_worksheet('Statistics')        # Defaults to Sheet

# Iterate over the data and write it out row by row.
for index in range(3):
    for i in range(len(Tests_List)):
        worksheet(index).write(row, col, Tests_List[i])
        col += 1


workbook.close()