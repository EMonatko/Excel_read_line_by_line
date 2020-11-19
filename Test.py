def main():
    print('Hi')
    today = datetime.now()
    #today = date.today()
    today = today.strftime('%y%y%m%d%H%M%S')
    print(today)
    wb = Workbook()
    ATP = wb.add_sheet('ATP')
    ESS = wb.add_sheet('ESS')
    Stat = wb.add_sheet('Statistics')
    Tests_List = ['Temp', 'SN', 'Output Power @ P1dBCP','Output Power Control Range/Resolution, FWD PWR Ind', 'Output IP3', 'LO Carrier Leakage','Sideband Suppression',
                  'Frequency Accuracy and Stability', 'A1 - Noise Figure vs. Gain', 'A1 - Gain variability', 'A1 - Image Suppression vs. Gain', 'Spurious',
                  'A2 - Noise Figure vs. Gain', 'A2 - Gain variability', 'A2 - Image Suppression vs. Gain', 'Average Power Consumption', 'Input Voltage', 'Digital Tests'
                  ]
    for index in range(len(Tests_List)):
        ATP.write(0, index, Tests_List[index])
        ESS.write(0, index, Tests_List[index])
        Stat.write(0, index, Tests_List[index])

    wb.save(f'{today}.xlsx')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    import os
    #import pandas as pd
    #import xlsxwriter
    import xlwt
    from xlwt import Workbook
    from datetime import datetime
    main()
