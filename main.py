from calculate import Calculate
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import math
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import xlwings as xw
from pathlib import Path
import openpyxl
PI = 3.14
KB = 1.38*(10**-23)
filepath = ''
sheet_input = ''
output_dir = ''
value_for_model = 100


def curve_fitting(df, value):
    def func(x, const1, const2):
        return const2*(np.exp(-const1 * x))

    tau_data = df.iloc[1:, :-1].values.flatten()
    g1_data = df.iloc[1:, -1].values.flatten()

    popt, pcov = curve_fit(func, tau_data, g1_data)
    const1, const2 = popt
    x_model = np.linspace(min(tau_data), max(tau_data), int(value))
    y_model = func(x_model, const1, const2)

    ax.set_xlabel("tau")
    ax.set_ylabel("g1(tau)")
    ax.plot(x_model, y_model, color='r')

    graph.draw()
    gamma_int.set(round( popt[0]/10**6 , 6)) # units: microsec^-1

    D = (popt[0] * 10**-14) / ((q ** 2) * 10**-6)
    d_int.set(round(D, 7)) #units: cm2/sec

    R = KB* (10 ** 13) * float(float(temperature_entry.get()) + 273) / (
            D  * 6 * PI * float(viscosity_entry.get()))
    r_int.set(round(R, 6)) #units: nm



def curve_plotting1():

    def curve_fit_func():
        ax.clear()
        ax.scatter(tau_data, g1_data, color='g')
        plt.xscale("log")
        graph.draw()

        value_label_1 = Label(text=value_for_model, width=8, font=14, bg="#CDE8E5", fg="black", anchor="center")
        value_label.grid(row=6, column=6)
        value_label_1.grid(row=6, column=6)
        curve_fitting(df, value_for_model)

    def increase_func():
        global value_for_model
        value_for_model += 100
        value_label_1 = Label(text=value_for_model, width=8, font=14, bg="#CDE8E5", fg="black", anchor="center")
        value_label_1.grid(row=6, column=6)

    def decrease_func():
        global value_for_model
        if value_for_model < 200:
            value_for_model = 100
        else:
            value_for_model -= 100
        value_label_1 = Label(text=value_for_model, width=8, font=14, pady=10, bg="#CDE8E5", fg="black", anchor="center")
        value_label_1.grid(row=6, column=6)

    df = pd.read_excel(filepath)

    df = df.dropna(how='all', axis=1)

    tau_data = df.iloc[1:, :-1].values
    g1_data = df.iloc[1:, -1].values

    ax.clear()
    ax.scatter(tau_data, g1_data, color='g')
    plt.xscale("log")
    graph.draw()

    extra_button_1 = Button(text='Curve Fit', width=15, bg="#4D869C", command=curve_fit_func)
    extra_button_1.config(padx=10, pady=10)
    extra_button_1.grid(row=5, column=5)

    plus_button = Button(text="+", fg='black', bg="#4D869C", padx=10, width=5, command=increase_func, pady=10, font=20)
    plus_button.grid(row=6, column=5)
    minus_button = Button(text="-", fg='black', bg="#4D869C", padx=10, width=5, command=decrease_func, pady=10, font=20)
    minus_button.grid(row=6, column=7)
    value_label = Label(text="100", width=8, font=14, bg="#CDE8E5", fg="black", anchor="center")
    value_label.grid(row=6, column=6)


def curve_plotting():

    def curve_fit_func():
        ax.clear()
        ax.scatter(tau_data, g1_data, color='g')
        plt.xscale("log")
        graph.draw()

        value_label_1 = Label(text=value_for_model, width=8, font=14, bg="#CDE8E5", fg="black", anchor="center")
        value_label_1.grid(row=6, column=6)
        curve_fitting(df, value_for_model)

    def increase_func():
        global value_for_model
        value_for_model += 100
        value_label_1 = Label(text=value_for_model, width=8, font=14, bg="#CDE8E5", fg="black", anchor="center")
        value_label_1.grid(row=6, column=6)

    def decrease_func():
        global value_for_model
        if value_for_model < 200:
            value_for_model = 100
        else:
            value_for_model -= 100
        value_label_1 = Label(text=value_for_model, width=8, font=14, bg="#CDE8E5", fg="black", anchor="center")
        value_label_1.grid(row=6, column=6)

    global sheet_input
    sheet_input = clicked.get()
    df = pd.read_excel(Path(output_dir) / f"{sheet_input}.xlsx")

    df = df.dropna(how='all', axis=1)

    tau_data = df.iloc[1:, :-1].values
    g1_data = df.iloc[1:, -1].values

    ax.clear()
    ax.scatter(tau_data, g1_data, color='g')
    plt.xscale("log")
    graph.draw()

    extra_button_1 = Button(text='Curve Fit', width=15, bg="#4D869C",command=curve_fit_func)
    extra_button_1.config(padx=10, pady=10)
    extra_button_1.grid(row=5, column=5)

    plus_button = Button(text="+", fg='black', bg="#4D869C", padx=10, width=5, command=increase_func, pady=10, font=20)
    plus_button.grid(row=6, column=5)
    minus_button = Button(text="-", fg='black', bg="#4D869C", padx=10, width=5, command=decrease_func, pady=10, font=20)
    minus_button.grid(row=6, column=7)
    value_label = Label(text="100", width=8, bg="#CDE8E5", fg="black", anchor="center", font=14)
    value_label.grid(row=6, column=6)


def graph_fit():

    file = Path(__file__).parent
    excel_file = Path(filepath)
    global output_dir
    output_dir = file / 'DLS_GUI_excel'

    output_dir.mkdir(parents=True, exist_ok=True)

    with xw.App(visible=False) as app:

        wb = app.books.open(excel_file)
        sheet_num = 0
        while sheet_num < 1:
            if len(wb.sheets) == 1:
                messagebox.showinfo(message='Please click on "Curve Plot"')
                curve_fit_button = Button(text='Curve Plot', width=15, padx=10, pady=10, bg="#4D869C", command=curve_plotting1)
                curve_fit_button.grid(row=1, column=4)

            else:
                messagebox.showinfo(
                    message='Your excel file contains more than one sheet. Enter the sheet name and click on "Curve Plot".')
                sheet_names = []
                for sheet in wb.sheets:
                    wb_new = app.books.add()
                    sheet.copy(after=wb_new.sheets[0])
                    wb_new.sheets[0].delete()
                    wb_new.save(output_dir / f'{sheet.name}.xlsx')
                    sheet_names.append(sheet.name)
                    wb_new.close()

                sheet_name = Label(text='Sheet Name', width=15, pady=10, bg="#CDE8E5", fg="black")
                sheet_name.grid(row=4, column=3)

                global clicked
                clicked = StringVar()
                clicked.set(sheet_names[0])
                drop = OptionMenu(window, clicked, *sheet_names)
                drop.grid(row=4, column=4,)
                curve_fit_button = Button(text='Curve Plot', width=15, padx=10, pady=10, bg="#4D869C", command=curve_plotting)
                curve_fit_button.grid(row=1, column=4)

            sheet_num = sheet_num + 1


def import_func():
    global filepath
    filepath = filedialog.askopenfilename()


def calculate_func():
    wave_length_input = wave_length_entry.get()
    refractive_index_input = refractive_index_entry.get()
    viscosity_input = viscosity_entry.get()
    scattering_angle_input = scattering_angle_entry.get()
    temperature_input = temperature_entry.get()

    calculate = Calculate(wave_length_input, refractive_index_input, viscosity_input, scattering_angle_input, temperature_input)
    calculate.file_path = filepath
    is_yes = calculate.check_input_values()
    is_ok = calculate.check_file_path()

    if is_ok == "YES" and is_yes != "NO":
        global q
        q = ( (4 * PI * float(refractive_index_input) ) * math.sin( float( scattering_angle_input )/114.592 ) ) / float(wave_length_input)
        q_int.set(str(round(q, 3))) # units: nm^-1
        graph_fit()


def quit_me():
    window.quit()
    window.destroy()


window = Tk()
window.protocol("WM_DELETE_WINDOW", quit_me)
window.title('DLS GUI')
window.configure(background='#CDE8E5')
window.config(width=500, height=500, padx=20, pady=20)
button = Button(text='Import Data', fg='black', bg="#4D869C", padx=10, width=15, command=import_func, pady=10)
button.grid(row=1, column=1)

wave_length = Label(text='Wave Length(nm)', width=15, pady=10, bg="#CDE8E5", fg="black")
wave_length_entry = Entry(width=15)
wave_length_entry.focus()
wave_length.grid(row=2, column=1)
wave_length_entry.grid(row=2, column=2)

refractive_index = Label(text='Refractive Index', width=15, pady=10, bg="#CDE8E5", fg="black")
refractive_index_entry = Entry(width=15)
refractive_index.grid(row=3, column=1)
refractive_index_entry.grid(row=3, column=2)

scattering_angle = Label(text='Scattering Angle\n(degrees)', width=15, bg="#CDE8E5", fg="black", pady=10)
scattering_angle_entry = Entry(width=15)
scattering_angle.grid(row=2, column=3)
scattering_angle_entry.grid(row=2, column=4)

viscosity = Label(text='Viscosity(Pa*s)', width=15, bg="#CDE8E5", fg="black", pady=10)
viscosity_entry = Entry(width=15)
viscosity.grid(row=3, column=3)
viscosity_entry.grid(row=3, column=4)

temperature = Label(text='Temperature(celsius)', width=15, bg="#CDE8E5", fg="black", pady=10)
temperature_entry = Entry(width=15)
temperature.grid(row=4, column=1)
temperature_entry.grid(row=4, column=2)

cal_button = Button(text='Calculate', fg="black", bg="#4D869C", width=15, command=calculate_func, padx=10, pady=10)
cal_button.grid(row=1, column=3)

q_int = StringVar()
cal_q = Label(text='Wave vector q:\n(nm^-1)', width=15, bg="#CDE8E5", fg="black", pady=10)
entry_q = Entry(width=15, textvariable=q_int)
cal_q.grid(row=5, column=1)
entry_q.grid(row=5, column=2)

gamma_int = StringVar()
gamma = Label(text='Gamma :\n(µs^-1)', width=15, bg="#CDE8E5", fg="black", pady=10)
gamma_entry = Entry(width=15, textvariable=gamma_int)
gamma.grid(row=6, column=1)
gamma_entry.grid(row=6, column=2)

d_int = StringVar()
d = Label(text='Diffusivity coefficient:\n(cm2/sec)', width=16, bg="#CDE8E5", fg="black", pady=10)
d_entry = Entry(width=15, textvariable=d_int)
d.grid(row=7, column=1)
d_entry.grid(row=7, column=2)

r_int = StringVar()
r = Label(text='Hydrodynamic Radius:\n(nm)', width=16, bg="#CDE8E5", fg="black", pady=10)
r_entry = Entry(width=15, textvariable=r_int)
r.grid(row=8, column=1)
r_entry.grid(row=8, column=2)

fig, ax = plt.subplots()
graph = FigureCanvasTkAgg(fig, master=window)
graph.get_tk_widget().grid(row=5, column=3, rowspan=4, columnspan=2, pady=10, padx=10)

toolbar = NavigationToolbar2Tk(graph, pack_toolbar=False)
toolbar.grid(row=9, column=3)

window.mainloop()