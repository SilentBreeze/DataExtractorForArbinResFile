from tkinter import filedialog
import pypyodbc
import numpy as np
import matplotlib.pyplot as plt


def getData(file):
    conn = pypyodbc.win_connect_mdb(file, readonly=True)
    cur = conn.cursor()
    cur.execute('SELECT test_time,current,voltage FROM Channel_Normal_Table')

    result = cur.fetchall()
    data = np.array(result)
    data = data[data[:, 0].argsort()]

    cur.close()
    conn.commit()
    conn.close()

    return data.tolist()


def call_back(event):
    axtemp = event.inaxes
    x_min, x_max = axtemp.get_xlim()
    fanwei = (x_max - x_min) / 10
    if event.button == 'up':
        axtemp.set(xlim=(x_min + fanwei, x_max - fanwei))
    elif event.button == 'down':
        axtemp.set(xlim=(x_min - fanwei, x_max + fanwei))
    fig.canvas.draw_idle()  # 绘图动作实时反映在图像上


filetypes = [("Arbin数据源文件", "*.res")]
filename = filedialog.askopenfilename(title='选择单个文件', filetypes=filetypes)
if len(filename) < 1:
    exit(0)

df = np.array(getData(filename))

fig = plt.figure()
fig.canvas.mpl_connect('scroll_event', call_back)
fig.canvas.mpl_connect('button_press_event', call_back)

ax1 = fig.add_subplot(111)
l1 = ax1.plot(df[:, 0], df[:, 2], linestyle='-',
              color='#103B59', label='Voltage')
ax2 = ax1.twinx()
l2 = ax2.plot(df[:, 0], df[:, 1], linestyle='-',
              color='#B55A0A', label='Current')

ls = l1+l2
labs = [l.get_label() for l in ls]
ax1.legend(ls, labs, loc=0)
ax1.grid()

ax1.set_xlabel('Time (s)')
ax1.set_ylabel('Current (A)')
ax2.set_ylabel('Voltage (V)')

plt.show()
