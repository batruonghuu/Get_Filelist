import datetime
import os
import time
import tkinter.filedialog
import xlwings as xw


ask_path = tkinter.filedialog.askdirectory(title="Select folder")
# Ask directory

os.chdir(ask_path)
# Active folder # Kích hoạt folder hiện hành

active_folder_name = os.getcwd()
# Get name of active folder

list_of_file = os.listdir()
# Get list of all file in active folder

new_file_excel = xw.Book()
# Open new Excel file

active_sheet = new_file_excel.sheets.active
# Get active sheet of Excel file

active_sheet.range("A3").value = "Folder" + active_folder_name
active_sheet.range("A4").value = "Count:" + str(len(list_of_file))
# Write the value in Book

row = 7
for name_file in list_of_file:

    path = os.path.join(active_folder_name,name_file)
    # Define the path
    # Remember character "r" before the path file

    mtime = os.path.getmtime(path)
    # Get modification time of a file

    m_time = datetime.datetime.fromtimestamp(mtime).strftime("%d-%m-%Y %H:%M:%S")
    # Convert timestamp format

    accesslast = os.path.getatime(path)
    # Get last access time of file

    access_last = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(accesslast))
    # Convert timestamp format

    size = os.path.getsize(path)
    # Get size of file

    daycreatetime = os.path.getctime(path)
    # Get day create time

    day_creat_time = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(daycreatetime))

    active_sheet.range("A" + str(row)).value = name_file
    active_sheet.range("B" + str(row)).value = path
    active_sheet.range("D" + str(row)).value = day_creat_time
    active_sheet.range("E" + str(row)).value = access_last
    active_sheet.range("G" + str(row)).value = round(size/1024,2)
    row = row + 1


