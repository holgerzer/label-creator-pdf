#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from label_creator import create_labels

class LabelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Label Generator")
        self.root.geometry("600x550")

        # Variables
        self.names_file_path = None
        self.selected_first_name = None
        self.selected_last_name = None
        self.title = None
        self.logo_1_path = None
        self.logo_2_path = None
        self.output_file_path = None

        # UI Elements
        self.label_select_file = tk.Label(root, text="Select an Excel file with names:")
        self.label_select_file.pack()

        self.button_select_file = tk.Button(root, text="Select File", command=self.select_names_file)
        self.button_select_file.pack(pady=10)

        self.label_title = tk.Label(root, text="Title (optional)")
        self.input_title = tk.Entry(root)

        self.first_name_label = tk.Label(root, text="Select First Name Column")
        self.first_name_combobox = ttk.Combobox(root)

        self.last_name_label = tk.Label(root, text="Select Last Name Column")
        self.last_name_combobox = ttk.Combobox(root)

        self.button_select_logo_1 = tk.Button(root, text="Select Logo 1", command=self.select_logo_1)
        self.button_select_logo_2 = tk.Button(root, text="Select Logo 2", command=self.select_logo_2)
        self.button_create = tk.Button(root, text="Create Labels as PDF", command=self.create_labels)
        self.button_open = tk.Button(root, text="Open PDF", command=self.open_pdf)

        # Hide elements initially
        self.hide_elements()

    def hide_elements(self):
        """ Hides UI elements that are only shown after file selection. """
        for widget in [self.label_title, self.input_title, self.first_name_label, self.first_name_combobox,
                       self.last_name_label, self.last_name_combobox, self.button_select_logo_1,
                       self.button_select_logo_2, self.button_create, self.button_open]:
            widget.pack_forget()

    def select_names_file(self):
        """ Opens file dialog and allows user to select an Excel file. """
        self.names_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if self.names_file_path:
            self.label_select_file.config(text=f"Selected file:\n{self.names_file_path}")
            df = pd.read_excel(self.names_file_path)
            self.update_dropdowns(df.columns.tolist())

    def update_dropdowns(self, columns):
        """ Populates dropdowns with column names from the Excel file. """
        self.first_name_combobox['values'] = columns
        self.last_name_combobox['values'] = columns

        if len(columns) >= 2:
            self.first_name_combobox.set(columns[0])
            self.last_name_combobox.set(columns[1])

        # Show UI elements for label customization
        for widget in [self.label_title, self.input_title, self.first_name_label, self.first_name_combobox,
                       self.last_name_label, self.last_name_combobox, self.button_select_logo_1,
                       self.button_select_logo_2, self.button_create]:
            widget.pack(pady=5)

    def select_logo_1(self):
        """ Select first logo file. """
        self.logo_1_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg;*.png")])

    def select_logo_2(self):
        """ Select second logo file. """
        self.logo_2_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg;*.png")])

    def create_labels(self):
        """ Calls label_creator to generate the labels PDF. """
        df = pd.read_excel(self.names_file_path, engine='openpyxl')

        names = list(df[[self.first_name_combobox.get(), self.last_name_combobox.get()]].itertuples(index=False, name=None))

        self.output_file_path = create_labels(
            names,
            self.input_title.get() if self.input_title.get().strip() else None,
            self.logo_1_path,
            self.logo_2_path
        )

        self.button_open.pack(pady=10)

    def open_pdf(self):
        """ Opens the generated PDF. """
        if self.output_file_path:
            os.startfile(self.output_file_path)