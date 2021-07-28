from openpyxl import Workbook
import os


book = Workbook()
sheet = book.active


def read_about_files(files_folders_list):
    for i in range(len(files_folders_list)):
        sheet[f'A{i + 1}'] = i + 1
        sheet[f'B{i + 1}'] = 'testtask'
        sheet[f'C{i + 1}'] = files_folders_list[i]
        file_type = ''
        file_name_list = list(files_folders_list[i])
        dot_index = -1
        dot = file_name_list[dot_index]
        while dot != '.':
            file_type += file_name_list[dot_index]
            dot_index -= 1
            if dot_index < -len(file_name_list):
                file_type = 'без расширения'
                break
            dot = file_name_list[dot_index]
        if file_type != 'без расширения':
            file_type = file_type[-1::-1]
        sheet[f'D{i + 1}'] = file_type


list_in_files = []
for dirpath, dirnames, filenames in os.walk("."):
    # перебрать каталоги
    for dirname in dirnames:
        # print("Каталог:", os.path.join(dirpath, dirname))
        list_in_files.append(os.path.join(dirpath, dirname))
    # перебрать файлы
    for filename in filenames:
        # print("Файл:", os.path.join(dirpath, filename))
        list_in_files.append(os.path.join(dirpath, filename))

# ['.\\test5', '.\\.DS_Store', '.\\my_script.py', '.\\test.txt', '.\\test2.rar', '.\\test3.dat.txt', '.\\test4', '.\\test_dir.py',
#  '.\\Текст задания.txt', '.\\test5\\test7', '.\\test5\\.DS_Store', '.\\test5\\test6.txt']

my_files = os.listdir()
read_about_files(my_files)

n = 1
for i in range(-1, -3, -1):
    list_file_dirname = list_in_files[i].split('\\')
    sheet[f'A{len(my_files) + n}'] = len(my_files) + n
    sheet[f'B{len(my_files) + n}'] = list_file_dirname[1]
    sheet[f'C{len(my_files) + n}'] = list_file_dirname[2]
    sheet[f'D{len(my_files) + n}'] = list_file_dirname[2].split('.')[-1]
    n += 1



book.save("sample.xlsx")