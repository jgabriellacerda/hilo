from tkinter import *
from tkinter.ttk import *
import os, sys
import xlrd
import xlwt
from xlutils.copy import copy
import random
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib import ticker
import datetime
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import numpy as np

root = Tk()
root.title("HiLo Clothes")
#root.geometry('1352x750+0+0')
root.resizable(False, False)

# Create style Object
style = Style()
style.configure('TButton', font=('calibri', 10, 'bold'), background="white", foreground="black")
style.map("TButton",
    foreground=[('pressed', 'pink'), ('active', 'pink')],
    background=[('pressed', '!disabled', 'pink'), ('active', 'white')]
    )
style.configure('TFrame', background="pink")
style.configure('TLabel', background="pink")


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


root.iconbitmap(resource_path("hilo_icon.ico"))

file_name = ("Vendas.xls")
workbook_r = xlrd.open_workbook(file_name)
sheet_r = workbook_r.sheet_by_index(0)
workbook_w = copy(workbook_r)
sheet_w = workbook_w.get_sheet(0)
#sheet_w = workbook_w.add_sheet(file_name, cell_overwrite_ok=True)

frame_lista = Frame(root, relief=RIDGE)
frame_lista.grid(row=0, column=0, sticky=W+E+N+S)

# for scrolling vertically
yscrollbar = Scrollbar(frame_lista)
yscrollbar.pack(side=RIGHT, fill=Y)

label = Label(frame_lista, text="Selecione clientes: ", font=("Times New Roman", 10))
label.pack(padx=5, pady=5)
list_box = Listbox(frame_lista, selectmode="multiple", yscrollcommand=yscrollbar.set)

# Widget expands horizontally and
# vertically by assigning both to
# fill option
list_box.pack(padx=5, pady=5, expand=YES, fill="both")

clients = []
data = []
dict_keys = []
dict_venda = {}
for i in range(sheet_r.nrows):
    if i > 0:
        novo_cliente = sheet_r.cell_value(i,1)
        if novo_cliente not in clients:
            clients.append(sheet_r.cell_value(i, 1))
    dict_venda = {}
    for j in range(sheet_r.ncols):
        if i == 0:
            dict_keys.append(sheet_r.cell_value(i,j))
        else:
            if j == 3:
                date = datetime.datetime(*xlrd.xldate_as_tuple(sheet_r.cell_value(i, j), sheet_r.book.datemode))
                dict_venda[dict_keys[j]] = date
            else:
                dict_venda[dict_keys[j]] = sheet_r.cell_value(i, j)
    if i == 0:
        print(dict_keys)
    else:
        data.append(dict_venda)
        print(dict_venda)

'''
for row in range(100):
    sheet_w.write(row+1, 1, clientes[random.randint(0, len(clientes)-1)])
    sheet_w.write(row + 1, 2, random.randint(30, 200))

workbook_w.save(file_name)
'''

for each_item in range(len(clients)):
    list_box.insert(END, clients[each_item])
    list_box.itemconfig(each_item, bg="pink")

# Attach listbox to vertical scrollbar
yscrollbar.config(command=list_box.yview)

frame_graph = Frame(root, relief=RIDGE)
frame_graph.grid(row=0, column=1, sticky=W+E)


fig = plt.figure(figsize=(5, 4), dpi=100)
ax = fig.add_subplot(111)
line = plt.plot([])
ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
locator = mdates.AutoDateLocator(minticks=4, maxticks=7)
fig.autofmt_xdate()
# Creating Canvas
canv = FigureCanvasTkAgg(fig, frame_graph)
get_widz = canv.get_tk_widget()
plt.grid(True)


def _plot_line():
    global line
    line_obj = line.pop()
    line_obj.remove()
    valores = []
    dates = []
    for row in data:
        valores.append(row['Valor'])
        dates.append(row['Data'])
    num_marks = 5
    date_max_num = max(mdates.date2num(dates))
    date_max = mdates.num2date(date_max_num)
    date_min_num = min(mdates.date2num(dates))
    date_min = mdates.num2date(date_min_num)
    locators_np = np.round(np.linspace(date_min_num, date_max_num, num_marks))
    locators = locators_np.tolist()
    ax.set_xlim([date_min, date_max])
    ax.set_ylim(0, max(valores)*1.05)
    locator = ticker.FixedLocator(locators)
    ax.xaxis.set_major_locator(locator)
    line = ax.plot(dates, valores, color='pink')
    canv.draw()
    get_widz.pack()


btn_plot = Button(frame_lista, text="Plotar", style='TButton', command=_plot_line)
btn_plot.pack(side=BOTTOM, padx=5, pady=5)


def _quit():
    root.quit()     # stops mainloop
    root.destroy()  # this is necessary on Windows to prevent
                    # Fatal Python Error: PyEval_RestoreThread: NULL tstate


frame_quit = Frame(root, relief=RIDGE)
frame_quit.grid(row=2, column=0, sticky=W+E+N+S, columnspan=2)
btn_quit = Button(frame_quit, text="Quit", command=_quit)
btn_quit.pack(side=BOTTOM, padx=5, pady=5)

root.mainloop()
