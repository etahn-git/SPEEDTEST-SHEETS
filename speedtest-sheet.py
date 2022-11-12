class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

print(bcolors.HEADER + bcolors.OKBLUE + bcolors.BOLD + '==================================================== SPEEDTEST SHEETS ====================================================' + bcolors.ENDC)


import speedtest as ss
import time
from pathlib import Path
from openpyxl import Workbook, load_workbook as excel
import sys

FILE_NAME = 'SHEET.xlsx'

path = Path(FILE_NAME)
if path.is_file():
    print(bcolors.FAIL + bcolors.BOLD + 'Please remove the old spreadsheet from the current directory.')
    sys.exit("")
else:
    print('')


st = ss.Speedtest()
local_time = time.localtime()
systemtime = time.strftime('%a, %d %b %Y %H:%M:%S', local_time)
date = '%b %d %Y'
time_ = '%H:%M:%S'
system_date = time.strftime(date, local_time)
system_time = time.strftime(time_, local_time)

wb = Workbook()
ws = wb.active
ws.title = "SPEEDTEST SHEET"


valid = False
while not valid:
    try:
        print(bcolors.OKGREEN + '------------------------------------------------ Ready! ------------------------------------------------' + bcolors.ENDC)
        print(bcolors.WARNING + 'How many x seconds should it wait to check & log again? !IN SECONDS!' + bcolors.ENDC)
        print(f'{bcolors.OKCYAN}Time converter: {bcolors.UNDERLINE}https://etahn.ml/time-converter' + bcolors.ENDC)
        looptime = int(input(bcolors.OKBLUE + bcolors.BOLD + 'Input: x= ' + bcolors.ENDC))
        valid = True
    except ValueError:
        print(bcolors.FAIL + 'Please only input digits' + bcolors.ENDC)
print(bcolors.OKGREEN + f'Started! Checking every {looptime} seconds.' + bcolors.ENDC)

while True:
    download_speed = st.download()
    upload_speed = st.upload()
    download = '{:5.2f} Mb'.format(download_speed/(1024*1024))
    upload = '{:5.2f} Mb'.format(upload_speed/(1024*1024))
    
    print(bcolors.OKCYAN + '-----------------------------------')
    print(systemtime)
    print('Download Speed: {:5.2f} Mb'.format(download_speed/(1024*1024) ))
    print('Upload Speed: {:5.2f} Mb'.format(upload_speed/(1024*1024) ))
    bcolors.ENDC
    time_ = '%H:%M:%S'

    local_time = time.localtime()
    systemtime = time.strftime('%a, %d %b %Y %H:%M:%S', local_time)
    date = '%b %d %Y, %a'
    time_ = '%H:%M:%S'
    system_date = time.strftime(date, local_time)
    system_time = time.strftime(time_, local_time)
    
    if system_date in wb.sheetnames:
        print('')
    else:
        wb.create_sheet(system_date)
        ws = wb[system_date]
        ws.append([system_date, 'Download / MB', 'Upload / MB'])
    
    ws.append([system_time, download, upload])
    wb.save(FILE_NAME)
    time.sleep(looptime)