from tkinter import *
from tkinter import messagebox
PI = 3.14


class Calculate:
    def __init__(self, WLI, RI, VI, SAI, TI):
        self.wave_length_input = WLI
        self.refractive_index_input = RI
        self.viscosity_input = VI
        self.scattering_angle_input = SAI
        self.temperature_input = TI
        self.file_path = ''
        self.q = 0

    def check_input_values(self):
        if self.wave_length_input == '' or self.refractive_index_input == '' or self.viscosity_input == '' or self.scattering_angle_input == '' or self.temperature_input == '':
            messagebox.showerror(message='Please fill all the entries!')
            return "NO"

        elif self.wave_length_input != '' or self.refractive_index_input != '' or self.viscosity_input != '' or self.scattering_angle_input != '' or self.temperature_input != '':
            try:
                float(self.wave_length_input)
                float(self.refractive_index_input)
                float(self.viscosity_input)
                float(self.scattering_angle_input)
                float(self.temperature_input)
            except ValueError:
                messagebox.showerror(message="Entered values must be numbers!!!")
                return "NO"

    def check_file_path(self):
        if self.file_path == '':
            messagebox.showerror(message='Please import the data file!')
            return "NO"
        else:
            return "YES"

