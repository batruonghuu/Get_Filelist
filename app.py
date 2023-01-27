import datetime
import os
import time

path = r"E:\Code Project\Learn Python\Filelist\doc\Truong Huu Ba.xlsx"
# Define the path
# Remember character "r" before the path file

m_time = os.path.getmtime(path)
# Get modification time of a file

dt_m = datetime.datetime.fromtimestamp(m_time).strftime("%d-%m-%Y %H:%M:%S")
# Convert timestamp format

accesslast = os.path.getatime(path)
# Get last access time of file

accesstime = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(accesslast))
# Convert timestamp format

size = os.path.getsize(path)
# Get size of file

print(m_time)
print(dt_m)
print(accesstime)
print(round(size/1024,2),"KB")
