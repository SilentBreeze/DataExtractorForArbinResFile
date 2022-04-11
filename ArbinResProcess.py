import gc
from tkinter.ttk import Progressbar
import numpy as np
from openpyxl import Workbook
import pypyodbc
import tkinter
from tkinter import Button, Frame, Label, StringVar, Tk, Toplevel, filedialog
from os.path import split as splitpath
from os.path import splitext
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.backend_bases import key_press_handler

OPEN_FILE_TYPE = [("Arbin数据源文件", "*.res")]
SAVE_FILE_TYPE = [("Excel工作簿", "*.xlsx")]
RES_TABLE_NAME = "Channel_Normal_Table"
RES_TABLE_FIELD = ["Data_Point", "Datetime", "Test_Time", "Step_Time", "Step_Index",
                   "Cycle_Index", "Current", "Voltage", "Charge_Capacity", "Discharge_Capacity", "Charge_Energy", "Discharge_Energy"]


def getResFile():
    filefullpath = filedialog.askopenfilename(
        title="选择单个文件", filetypes=OPEN_FILE_TYPE)

    if len(filefullpath) < 1:
        return ('', '', '', '')

    (filepath, filefullname) = splitpath(filefullpath)
    (filename, fileext) = splitext(filefullname)

    return (filefullpath, filepath, filename, fileext)


def getResFiles():
    filefullpaths = filedialog.askopenfilenames(
        title="选择文件", filetypes=OPEN_FILE_TYPE)

    if len(filefullpaths) < 1:
        return []

    fileinfos = list()
    for filefullpath in filefullpaths:
        (filepath, filefullname) = splitpath(filefullpath)
        (filename, fileext) = splitext(filefullname)
        fileinfos.append((filefullpath, filepath, filename, fileext))

    return fileinfos


def selectSavePath():
    path = filedialog.askdirectory(title='选择存放目录')

    if len(path) == 0:
        return ''

    return path


def loadData(file):
    fieldname_str = ','.join(field.lower() for field in RES_TABLE_FIELD)
    execute_str = 'SELECT %s FROM %s' % (fieldname_str, RES_TABLE_NAME)

    conn = pypyodbc.win_connect_mdb(file, readonly=True)
    cur = conn.cursor()
    cur.execute(execute_str)

    result = cur.fetchall()

    cur.close()
    conn.close()

    del conn, cur
    gc.collect()

    data = np.array(result)

    return data[data[:, 0].argsort()]


def saveData(filepath, filename, data):
    wb = Workbook()
    ws = wb.active
    ws.append(RES_TABLE_FIELD)
    for row in data:
        ws.append(row)
    wb.save(filepath+"/"+filename+".xlsx")
    wb.close()


def plotData(filename, data):
    window = Toplevel(main)
    window.title(filename)
    window.focus_force()

    frame1 = Frame(window)
    frame2 = Frame(window)

    fig = Figure(figsize=(8, 6), dpi=100)
    ax1 = fig.add_subplot(111)
    l1 = ax1.plot(data[:, 2], data[:, 6], linestyle='-',
                  color='#B55A0A', label='Current')
    ax2 = ax1.twinx()
    l2 = ax2.plot(data[:, 2], data[:, 7], linestyle='-',
                  color='#103B59', label='Voltage')

    ls = l1+l2
    labs = [l.get_label() for l in ls]
    ax1.legend(ls, labs, loc=0)
    ax1.grid()

    ax1.set_xlabel('Time (s)')
    ax1.set_ylabel('Current (A)')
    ax2.set_ylabel('Voltage (V)')

    canvas = FigureCanvasTkAgg(fig, master=frame1)
    canvas.draw()
    toolbar = NavigationToolbar2Tk(canvas, frame1)
    toolbar.update()
    canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    labelText = StringVar()
    labelText.set("")
    saveIndicator = Label(frame2, textvariable=labelText)
    saveIndicator.pack(side=tkinter.LEFT, fill=tkinter.BOTH, anchor=tkinter.W)

    def save_command():
        labelText.set("")
        window.update()
        fullsavefilename = filedialog.asksaveasfilename(
            title="导出数据文件", filetypes=SAVE_FILE_TYPE, initialfile=filename)

        (savefilepath, fullfilename) = splitpath(fullsavefilename)
        (savefilename, savefileext) = splitext(fullfilename)
        window.focus_force()

        if len(fullsavefilename) == 0:
            return

        labelText.set("数据正在导出......")
        window.update()
        saveData(savefilepath, savefilename, data.tolist())
        labelText.set("数据已导出为："+savefilepath+"/"+savefilename+".xlsx")

    Button(frame2, text="导出数据", command=save_command).pack(
        side=tkinter.RIGHT, fill=tkinter.Y, anchor=tkinter.W)

    frame1.pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)
    frame2.pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    def call_back(event):
        axtemp1 = event.canvas.figure.axes[0]
        axtemp2 = event.canvas.figure.axes[1]
        x_min, x_max = axtemp1.get_xlim()
        y_min_1, y_max_1 = axtemp1.get_ylim()
        y_min_2, y_max_2 = axtemp2.get_ylim()
        x_fanwei = (x_max - x_min) / 10
        y_fanwei_1 = (y_max_1 - y_min_1) / 10
        y_fanwei_2 = (y_max_2 - y_min_2) / 10
        if event.button == 'up':
            axtemp1.set(xlim=(x_min + x_fanwei, x_max - x_fanwei))
            axtemp1.set(ylim=(y_min_1 + y_fanwei_1, y_max_1 - y_fanwei_1))
            axtemp2.set(ylim=(y_min_2 + y_fanwei_2, y_max_2 - y_fanwei_2))
        elif event.button == 'down':
            axtemp1.set(xlim=(x_min - x_fanwei, x_max + x_fanwei))
            axtemp1.set(ylim=(y_min_1 - y_fanwei_1, y_max_1 + y_fanwei_1))
            axtemp2.set(ylim=(y_min_2 - y_fanwei_2, y_max_2 + y_fanwei_2))
        canvas.draw_idle()

    def on_key_press(event):
        print("you pressed {}".format(event.key))
        key_press_handler(event, canvas, toolbar)

    canvas.mpl_connect('scroll_event', call_back)
    canvas.mpl_connect('button_press_event', call_back)
    canvas.mpl_connect("key_press_event", on_key_press)


def plot_command():
    (filefullpath, filepath, filename, fileext) = getResFile()

    if len(filefullpath) == 0:
        return

    data = loadData(filefullpath)
    plotData(filename, data)

    gc.collect()


def set_in_windows_center(tk_widget):
    root = tk_widget.master

    r_width = root.winfo_width()
    r_height = root.winfo_height()

    r_x = root.winfo_rootx()
    r_y = root.winfo_rooty()

    width = int(r_width/3*2)
    height = int(r_height/3*1)

    x = r_x + int(r_width/3/2)
    y = r_y + int(r_height/3/1)

    tk_widget.geometry("%sx%s+%s+%s" % (width, height, x, y))


def extract_command():
    main.attributes("-disabled", 1)

    path_list = getResFiles()
    path_num = len(path_list)
    if path_num == 0:
        main.attributes("-disabled", 0)
        return

    save_path = selectSavePath()
    if len(save_path) == 0:
        main.attributes("-disabled", 0)
        return

    window = Toplevel()
    window.title("正在导出")
    window.focus_force()
    set_in_windows_center(window)
    window.overrideredirect(True)
    main.wm_attributes('-topmost', 1)
    window.wm_attributes('-topmost', 1)

    pb = Progressbar(window, maximum=path_num+0.1,
                     mode="determinate", orient=tkinter.HORIZONTAL)
    pb.pack(side=tkinter.TOP, anchor=tkinter.CENTER,
            fill=tkinter.BOTH, expand=True)

    pb["value"] += 0.1
    window.update()
    for (filefullpath, filepath, filename, fileext) in path_list:
        data = loadData(filefullpath)
        saveData(save_path, filename, data.tolist())
        pb["value"] += 1
        window.update()
    window.destroy()
    main.wm_attributes('-topmost', 0)
    main.attributes("-disabled", 0)

    gc.collect()


main = Tk()
main.title("Arbin数据源文件处理")
main.geometry("300x100")
main.resizable(0,0)

Button(main, text="数据绘图", command=plot_command).pack(
    padx=10, pady=10, side=tkinter.TOP, fill=tkinter.BOTH)
Button(main, text="数据导出", command=extract_command).pack(
    padx=10, pady=10, side=tkinter.TOP, fill=tkinter.BOTH)

main.mainloop()
