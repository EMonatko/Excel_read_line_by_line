import os

def show_all_folders(path, change_value):
    os.chdir(path)

    d = '.'
    folders_list = [os.path.join(d, o) for o in os.listdir(d) if os.path.isdir(os.path.join(d, o))]
    for folder_len in range(len(folders_list)):
        folders_list[folder_len] = folders_list[folder_len].replace('.\\', '')
    print(folders_list, '\n')
    for folder_len in range(len(folders_list)):
        temp_path = os.path.join(path, folders_list[folder_len])
        print(f'Entering into {folders_list[folder_len]} Folder\nIn {temp_path}')
        try:
            excel_files_list = sorted([f for f in os.listdir(temp_path) if f.endswith('.xlsx')])
            print(excel_files_list)
        except FileNotFoundError:
            print(f'\nNo excel files in {folders_list[folder_len]}\n')
            excel_files_list = []

        if len(excel_files_list) != 0:
            for excel_files_in_folder in range(0, len(excel_files_list)):
                if change_value == excel_files_list[excel_files_in_folder]:
                    try:
                        os.rename(str(os.path.join(temp_path, excel_files_list[excel_files_in_folder])), (os.path.join(temp_path, f'{folders_list[folder_len]}.xlsx')))
                    except FileNotFoundError:
                        print(f'\nNo excel files in {folders_list[folder_len]}\n')

#os.rename(r'C:\Users\Ron\Desktop\Test\Products.txt',r'C:\Users\Ron\Desktop\Test\Shipped Products.txt')






def main(name) :
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
    path_string = 'D:\Rasberry Pie\Python'
    which_value_to_change_string = 'test.xlsx'
    show_all_folders(path_string, which_value_to_change_string)


# Press the green button in the gutter to run the script.
if __name__ == '__main__' :

    main('Rename Excel files to there folder name')