import os
import xlsxwriter
import speedtest
import schedule
import time

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

from datetime import datetime
from threading import Timer


workbook_name = 'speedtest.xlsx'
worksheet_name = 'Speed Test'


"""
"""

def workbook_not_exists():
    return workbook_name not in os.listdir()


"""
"""

def create_workbook():
    print('Creating Workbook...')

    workbook = xlsxwriter.Workbook(workbook_name)
    worksheet = workbook.add_worksheet(worksheet_name)

    worksheet.write('A1', 'Date')
    worksheet.write('B1', 'Time')
    worksheet.write('C1', 'Download (ms)')
    worksheet.write('D1', 'Upload (ms)')
    worksheet.write('E1', 'Ping (ms)')
    worksheet.write('F1', 'Server')
    worksheet.write('G1', 'IP Address')
    worksheet.write('H1', 'Results')

    workbook.close()


"""
"""

def plot_results(sheet_appended):
    plt.title('Speed Test')
    plt.xlabel('Test Number')

    plt.grid()
    sns.lineplot(data=sheet_appended, x=sheet_appended.index, y="Download (ms)")

    plt.show()
    #plt.savefig(fname='results.png')


"""
"""

def speed_test():
    speed_sheet = pd.read_excel(workbook_name)

    date = datetime.now().strftime('%d/%m/%Y')
    time = datetime.now().strftime('%H:%M:%S')

    print('Testing Speed...')

    speed = speedtest.Speedtest()

    speed.get_best_server()
    speed.download(threads=None)
    speed.upload(threads=None)
    speed.results.share()

    results = speed.results.dict()

    download = round(results['download'] * (10**-6), 2)
    upload = round(results['upload'] * (10**-6), 2)
    ping = round(results['ping'], 2)
    server = results['server']['sponsor']
    ip = results['client']['ip']
    png = results['share']

    sheet_appended = speed_sheet.append(pd.DataFrame(
        [[date, time, download, upload, ping, server, ip, png]], 
        columns=speed_sheet.columns),
        ignore_index=True)

    print(f'\nResults: \n\n{sheet_appended.tail()}')

    sheet_appended.to_excel(workbook_name, sheet_name=worksheet_name, index=False)

    print('\nSpeed Test Finalised and Worksheet Succesfully Updated!')

    plot_results(sheet_appended)


"""
"""


if __name__ == '__main__':
    if workbook_not_exists():
        create_workbook()
    
    schedule.every(1800).seconds.do(speed_test)
    
    while True:
        schedule.run_pending()
        time.sleep(1)
