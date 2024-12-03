import tkinter as tk
from tkinter import ttk


class Combobox:
    combox_list = {}

    def __init__(self, master, name, values, *args):
        self.combobox = ttk.Combobox(master, values=values, width=25)
        self.combobox.place(x=args[0], y=args[1])
        # self.combobox.current(0)
        self.combox_list[name] = self.combobox

    def get_box(self):
        return self.combobox

    def bind(self, action, func):
        self.combobox.bind(action, func)


class Checkbuttons:
    checkbutton_list = {}

    def __init__(self, master, name, variable, *args):
        self.checkbutton = tk.Checkbutton(master, text=name, variable=variable, offvalue=0, onvalue=1)
        self.checkbutton.place(x=args[0], y=args[1])
        self.checkbutton_list[name] = [self.checkbutton, variable]


class MyButton:
    button_list = []

    def __init__(self, master, text, command, *args):
        self.button = tk.Button(master, text=text, command=command, height=1)
        self.button.place(x=args[0], y=args[1])
        self.button_list.append(self.button)


class MyEntry:
    entry_list = {}

    def __init__(self, master, label_name, text, *args):
        self.entries = tk.Entry(master, width=15)
        self.entries.place(x=args[0], y=args[1])
        self.entries.insert(0, text)
        self.entry_list[label_name] = self.entries


class MyLabel:
    label_list = {}

    def __init__(self, master, label_name, text, *args):
        self.label = tk.Label(master, text=text)
        self.label.place(x=args[0], y=args[1])
        self.label_list[label_name] = self.label
        # print(self.label_list)

    def set_text(self, name, text):
        self.label_list[name].config(text=text)


class Settings:
    setting_obj = []

    def __init__(self, file_path):
        self.file_path = file_path
        self.settings = {}
        self.load_settings()

    def load_settings(self):
        with open(self.file_path, 'r') as file:
            section = None
            for line in file:
                line = line.strip()
                if line.startswith('[') and line.endswith(']'):
                    section = line[1:-1]
                    self.settings[section] = {}
                elif '=' in line:
                    key, value = line.split('=', 1)
                    if section:
                        self.settings[section][key.strip()] = (value.strip())

    def save_settings(self, param_dict):
        with open(self.file_path, 'w') as file:
            file.write('[Temperature]\n')
            for key, value in param_dict.items():
                file.write(f'{key}={value}\n')

            file.write('\n')
        # print(self.settings)
