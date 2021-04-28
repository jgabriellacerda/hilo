from __future__ import print_function
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.ttk import *
import os
import sys
import xlrd
import xlwt
from xlutils.copy import copy
import random
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib import ticker
from matplotlib import colors
import datetime as dt
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import numpy as np
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
import seaborn as sns

root = Tk()
root.title("HiLo Clothes")
root.geometry('1280x720+0+0')
#root.resizable(False, False)

data = []

frame_menu = Frame(root, relief=RIDGE)
frame_menu.pack(side=TOP, fill=X)

frame_content = Frame(root)
frame_content.pack(side=TOP, fill=BOTH, expand=True)
#frame_content.rowconfigure(0, weight=1)
#frame_content.columnconfigure(0, weight=1)
#frame_content.grid_propagate(0)

frame_footer = Frame(root, relief=RIDGE)
frame_footer.pack(side=BOTTOM, fill=X)

root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

frame_indicadores = Frame(frame_content)
frame_indicadores.grid(row=0, column=0, sticky=NSEW)

frame_config = Frame(frame_content)
frame_config.grid(row=0, column=0, sticky=NSEW)


def ds_selected(self):
    selected_data_source = data_sources.get()
    if selected_data_source == 'Arquivo do Excel':
        file_name = askopenfilename()
        if file_name:
            ds_entry.delete(0, END)
            ds_entry.insert(0, file_name)
    if selected_data_source == 'URL do GoogleSheets':
        file_name = 'https://docs.google.com/spreadsheets/d/1ysgqGzSqx3oyM1vzmeDWJ2_xxWntEexvPevY6BCRiRM/edit#gid=1275477312'
        if file_name:
            ds_entry.delete(0, END)
            ds_entry.insert(0, file_name)


label_select_ds = Label(frame_config, text='Selecione a fonte de dados:')
label_select_ds.pack(padx=5, pady=5)
vlist = ['Arquivo do Excel', 'URL do GoogleSheets']
data_sources = Combobox(frame_config, values=vlist)
data_sources.set("-")
data_sources.bind('<<ComboboxSelected>>', ds_selected)
data_sources.pack(padx=5, pady=5)
ds_entry = Entry(frame_config)
ds_entry.pack(padx=5, pady=5)


def read_data():
    selected_data_source = data_sources.get()
    if selected_data_source == 'Arquivo do Excel':
        file_name = ds_entry.get()
        read_ok = read_xls_file(file_name)
        if read_ok:
            label_state['text'] = 'Arquivo carregado.'
            btn_grafico['state'] = 'normal'
            clients = read_clients_names(data)
            insert_client_list(clients)
    elif selected_data_source == 'URL do GoogleSheets':
        url = ds_entry.get()
        url_split = url.split('/')
        print(url_split)


label_state = Label(frame_config, text='')
label_state.pack(side=BOTTOM, padx=5, pady=5)
btn_read_data = Button(frame_config, text="Carregar dados", command=read_data)
btn_read_data.pack(side=BOTTOM, padx=5, pady=5)


def update_content(content):
    frame_grafico.grid_remove()
    frame_seaborn.grid_remove()
    frame_indicadores.grid_remove()
    frame_config.grid_remove()
    if content == 'grafico':
        frame_grafico.grid()
    elif content == 'seaborn':
        frame_seaborn.grid()
    elif content == 'indicadores':
        frame_indicadores.grid()
    elif content == 'config':
        frame_config.grid()


def _show_grafico():
    update_content('grafico')


def _show_indicadores():
    update_content('indicadores')


def _show_seaborn():
    update_content('seaborn')

def _show_config():
    update_content('config')


btn_grafico = Button(frame_menu, text="Gráfico", command=_show_grafico, state='disabled')
btn_grafico.pack(side=LEFT, padx=5, pady=5)
btn_indicadores = Button(frame_menu, text="Indicadores", command=_show_indicadores, state='disabled')
btn_indicadores.pack(side=LEFT, padx=5, pady=5)
btn_seaborn = Button(frame_menu, text="Seaborn", command=_show_seaborn, state='disabled')
btn_seaborn.pack(side=LEFT, padx=5, pady=5)
btn_config = Button(frame_menu, text="Config", command=_show_config)
btn_config.pack(side=RIGHT, padx=5, pady=5)

def config_style():
    # Create style Object
    style_ = Style()
    style_.configure('TButton', font=('calibri', 10, 'bold'), background="white", foreground="black")
    style_.map("TButton",
               foreground=[('pressed', 'pink'), ('active', 'pink')],
               background=[('pressed', '!disabled', 'pink'), ('active', 'white')]
               )
    style_.configure('TFrame', background="pink")
    style_.configure('TLabel', background="pink")
    return style_


style = config_style()


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


root.iconbitmap(resource_path("hilo_icon.ico"))

#file_name = "Vendas.xls"
def read_xls_file(file_name):
    if file_name:
        workbook_r = xlrd.open_workbook(file_name)
        sheet_r = workbook_r.sheet_by_index(0)
        workbook_w = copy(workbook_r)
        sheet_w = workbook_w.get_sheet(0)

        dict_keys = read_dict_keys(sheet_r)
        global data
        data = read_sells_data(sheet_r, dict_keys)
        return True
    else:
        return False


frame_grafico = Frame(frame_content)
frame_grafico.grid(row=0, column=0, sticky=NSEW)
frame_content.rowconfigure(0, weight=1)
frame_content.columnconfigure(0, weight=1)


frame_lista = Frame(frame_grafico)
frame_lista.pack(side=LEFT, fill=Y)

# for scrolling vertically
yscrollbar = Scrollbar(frame_lista)
yscrollbar.pack(side=RIGHT, fill=Y)

label = Label(frame_lista, text="Selecione clientes: ", font=("Times New Roman", 10))
label.pack(padx=5, pady=5)
list_box = Listbox(frame_lista, selectmode="multiple", yscrollcommand=yscrollbar.set)

# Widget expands horizontally and
# vertically by assigning both to
# fill option
list_box.pack(padx=5, pady=5, expand=YES, fill=Y)


def read_dict_keys(sheet_r):
    dict_keys_ = []
    for j in range(sheet_r.ncols):
        dict_keys_.append(sheet_r.cell_value(0, j))
    return dict_keys_


def read_sells_data(sheet_r, dict_keys):
    data_ = []
    for i in range(1, sheet_r.nrows):
        dict_sell = {}
        for j in range(sheet_r.ncols):
            if j == 3:
                date = dt.datetime(*xlrd.xldate_as_tuple(sheet_r.cell_value(i, j), sheet_r.book.datemode))
                dict_sell[dict_keys[j]] = date
            else:
                dict_sell[dict_keys[j]] = sheet_r.cell_value(i, j)
        data_.append(dict_sell)
    return data_


def read_clients_names(data_):
    clients_ = []
    for i in range(0, len(data_) - 1):
        new_client = data_[i]['Nome']
        if new_client not in clients_:
            clients_.append(new_client)
    return clients_


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
#SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SPREADSHEET_ID = '1ysgqGzSqx3oyM1vzmeDWJ2_xxWntEexvPevY6BCRiRM'
RANGE_NAME = 'Saida/2020!A1:D'


def create_service():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    data_frame_ = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                resource_path('credentials.json'), SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    return service


def read_google_sheet(service, data_range, type):

    # Create drive service
    #drive_service = build('drive', 'v3', credentials=creds)
    #drive_service.files().copy(fileId=SAMPLE_SPREADSHEET_ID, body={"mimeType" = "application/vnd.google-apps.spreadsheet"}).execute()

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=data_range).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        data_frame_ = pd.DataFrame(values[1:], columns=values[0])
        data_frame_['Valor'] = data_frame_['Valor'].apply(lambda x: x.replace('R$', '').replace('.', '').replace(',', '.').replace(' ', '')
        if isinstance(x, str) else x).astype(float)
        for i, v in data_frame_['Data'].items():
            if v == "":
                data_frame_['Data'][i] = data_frame_['Data'][i-1]

        data_frame_['Data'] = data_frame_['Data'].replace(r'^\s*$', np.nan, regex=True)
        data_frame_['Data'] = data_frame_['Data'].apply(lambda x: pd.to_datetime(x, format='%d/%m/%Y'))
        dim = data_frame_.shape
        rows = dim[0]
        types = [type]*rows
        data_frame_["Tipo"] = types
        data_frame_['MesAno'] = data_frame_['Data'].dt.strftime('%m/%Y')

    return data_frame_


service = create_service()
#data_range = 'Saida/2019!A1:D'
#saida_2019 = read_google_sheet(service, data_range, 'saida')
data_range = 'Saida/2020!A1:D'
saida_2020 = read_google_sheet(service, data_range, 'saida')
saida2020_bymonth = saida_2020.groupby('MesAno').sum().reset_index(level=['MesAno'])
saida2020_bymonth['Tipo'] = 'saida'
print('saida=',saida2020_bymonth)

#data_range = 'Saida/2021*!A1:D'
#data_frame2021 = read_google_sheet(service, data_range, 'saida')
data_range = 'Entrada/2020!A2:G'
entrada_2020 = read_google_sheet(service, data_range, 'entrada')
entrada2020_bymonth = entrada_2020.groupby('MesAno').sum().reset_index(level=['MesAno'])
entrada2020_bymonth['Tipo'] = 'entrada'
print(entrada2020_bymonth)

data_frame2020 = pd.concat([saida2020_bymonth, entrada2020_bymonth], join='inner')
data_frame2020.loc[data_frame2020['Tipo'] == 'saida', ['Valor']] *= -1
diff2020 = data_frame2020.groupby('MesAno').sum().reset_index(level='MesAno')
diff2020['Tipo'] = 'diferenca'
data_frame2020 = pd.concat([data_frame2020, diff2020], join='inner')
data_frame2020.loc[data_frame2020['Tipo'] == 'saida', ['Valor']] *= -1
print(data_frame2020)

'''
for row in range(100):
    sheet_w.write(row+1, 1, clientes[random.randint(0, len(clientes)-1)])
    sheet_w.write(row + 1, 2, random.randint(30, 200))

workbook_w.save(file_name)
'''


def insert_client_list(clients):
    for each_item in range(len(clients)):
        list_box.insert(END, clients[each_item])
        list_box.itemconfig(each_item, bg="pink")


# Attach listbox to vertical scrollbar
yscrollbar.config(command=list_box.yview)

frame_graph = Frame(frame_grafico, relief=RIDGE)
frame_graph.pack(side=LEFT, anchor=CENTER, fill=BOTH, expand=True)

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
    line_obj = line[0]
    print(line_obj)
    try:
        line_obj.remove()
    except Exception:
        line_obj = line.pop()
        line_obj.remove()
    selected_clients = list_box.curselection()
    clients_str = []
    clients = read_clients_names(data)
    for client_num in selected_clients:
        clients_str.append(clients[client_num])
    valores = []
    dates = []
    for row in data:
        if row['Nome'] in clients_str:
            valores.append(row['Valor'])
            dates.append(row['Data'])
    num_marks = 5
    date_max_num = max(mdates.date2num(dates))
    date_max = mdates.num2date(date_max_num)
    date_min_num = min(mdates.date2num(dates))
    date_min = mdates.num2date(date_min_num)
    plot_dates = np.arange(date_min_num, date_max_num, 1)
    dates_num = (mdates.date2num(dates)).tolist()
    values = []
    for indice in range(len(plot_dates)):
        if plot_dates[indice] in dates_num:
            value_indice = dates_num.index(plot_dates[indice])
            values.append(valores[value_indice])
        else:
            values.append(0)
    locators_np = np.round(np.linspace(date_min_num, date_max_num, num_marks))
    locators = locators_np.tolist()
    ax.set_xlim([date_min, date_max])
    ax.set_ylim(0, max(valores) * 1.05)
    locator_ = ticker.FixedLocator(locators)
    ax.xaxis.set_major_locator(locator_)
    line = ax.bar(plot_dates, values, 1, color='pink')
    canv.draw()
    get_widz.pack(anchor=CENTER, expand=True, fill=BOTH)


btn_plot = Button(frame_lista, text="Plotar", style='TButton', command=_plot_line)
btn_plot.pack(side=BOTTOM, padx=5, pady=5)


frame_seaborn = Frame(frame_content)
frame_seaborn.grid(row=0, column=0, sticky=NSEW)
frame_seaborn.rowconfigure(0, weight=1)
frame_seaborn.columnconfigure(0, weight=1)


custom_palette = ["green", "red", "blue", "orange", "yellow", "purple"]
sns.set_palette(custom_palette)
sns.set_style("whitegrid")


def update_plot():
    facet = sns.catplot(x='MesAno', y='Valor', hue='Tipo', data=data_frame2020, kind='bar', ci=None, hue_order=['entrada', 'saida', 'diferenca'])
    seaborn_canvas.insert_figure(facet)


btn_plot = Button(frame_seaborn, text="Plotar", style='TButton', command=update_plot)
btn_plot.grid(row=1, column=0, sticky=S)


class SnsCanvas:

    def __init__(self, tkinter_frame):
        self.frame = tkinter_frame
        self.figure = plt.figure()
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.frame)
        self.canvas.draw()
        self.widget = self.canvas.get_tk_widget()
        self.widget.grid(row=0, column=0, sticky=NSEW)

    def insert_figure(self, facet):
        for item in self.canvas.get_tk_widget().find_all():
            self.canvas.get_tk_widget().delete(item)
        self.canvas.get_tk_widget().destroy()
        plt.close(self.figure.clear())
        self.figure = facet.fig
        self.canvas = FigureCanvasTkAgg(self.figure, master=frame_seaborn)
        facet.axes[0][0].yaxis.set_major_formatter(ticker.StrMethodFormatter('R${x:1.2f}'))
        ax.yaxis.set_major_formatter(ticker.PercentFormatter(1))
        self.canvas.draw()
        self.widget = self.canvas.get_tk_widget()
        self.widget.grid(row=0, column=0, sticky=NSEW)


seaborn_canvas = SnsCanvas(frame_seaborn)


def _quit():
    root.quit()  # stops mainloop
    root.destroy()  # this is necessary on Windows to prevent
    # Fatal Python Error: PyEval_RestoreThread: NULL tstate


btn_quit = Button(frame_footer, text="Quit", command=_quit)
btn_quit.pack(side=BOTTOM, padx=5, pady=5)

update_content('config')


if __name__ == '__main__':
    root.mainloop()
