import os
from openpyxl import Workbook
import time

'''Можете запускать программу сколько угодно раз.
И не надо каждый раз придумывать имя для файла.
Она генерируется из текущего времени и добавлением к нему своих слов'''


def create_files_folders_name_list(all_files_folders):
    """записывает все названия файлов и директории в список
    и возвращает этот список"""
    for dirpath, dirnames, filenames in os.walk("."):
        # перебрать каталоги
        for dirname in dirnames:
            all_files_folders.append(os.path.join('кат', dirpath, dirname))
        # перебрать файлы
        for filename in filenames:
            all_files_folders.append(os.path.join(dirpath, filename))
    return all_files_folders


def write_in_excel_file(all_files_folders):
    '''функция для записи в эксель файл. принимает
    в аргументы спискок файлов и директории'''
    n = 0
    for x in all_files_folders:
        n += 1
        if x.startswith('кат'):
            n -= 1
            continue
        else:
            folder_file_list = x.split('\\')
            sheet[f'A{n}'] = n
            if len(folder_file_list) == 2:
                sheet[f'B{n}'] = 'testtask'
            else:
                sheet[f'B{n}'] = folder_file_list[-2]
            sheet[f'C{n}'] = folder_file_list[-1]
            find_file_type = folder_file_list[-1].split('.')[-1]
            sheet[f'D{n}'] = find_file_type


if __name__ == '__main__':
    book = Workbook()
    sheet = book.active
    my_files_folders_list = []
    changing_files_folders_list = create_files_folders_name_list(my_files_folders_list)
    write_in_excel_file(changing_files_folders_list)
    book.save(f"{time.strftime('%H%M%S')}_all_file_names.xlsx")