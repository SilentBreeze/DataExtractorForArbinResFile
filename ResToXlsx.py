import os
from tkinter import filedialog
import pypyodbc
import numpy as np
from openpyxl import Workbook


def getData(file):
    conn = pypyodbc.win_connect_mdb(file, readonly=True)
    cur = conn.cursor()
    cur.execute('SELECT data_point,test_time,step_time,step_index,cycle_index,current,voltage,charge_capacity,discharge_capacity FROM Channel_Normal_Table')

    result = cur.fetchall()
    data = np.array(result)
    data = data[data[:, 0].argsort()]

    cur.close()
    conn.commit()
    conn.close()

    return data.tolist()


def writeData(savepath, filename, data):
    wb = Workbook()
    ws = wb.active
    ws.append(['Data_Point', 'Test_Time', 'Step_Time', 'Step_Index', 'Cycle_Index',
               'Current', 'Voltage', 'Charge_Capacity', 'Discharge_Capacity'])
    for row in data:
        ws.append(row)
    wb.save(f'{savepath}/{filename}.xlsx')


filetypes = [("Arbin数据源文件", "*.res")]
filenames = filedialog.askopenfilenames(title='选择文件', filetypes=filetypes)

save_path = filedialog.askdirectory(title='选择存放目录')
if len(save_path) < 1:
    save_path = '.'

for file in filenames:
    (file_path, fullfilename) = os.path.split(file)
    (file_name, file_ext) = os.path.splitext(fullfilename)

    writeData(save_path, file_name, getData(file))
