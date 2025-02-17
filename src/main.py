# pip install pandas reportlab openpyxl
# pyinstaller --onefile --add-data "stag-sans-thin.ttf;." --add-data "stag-sans-book.ttf;." --add-data "stag-sans-light-italic.ttf;." --add-data "stag-sans-medium.ttf;." label_print.py

import os
import sys
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk


def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):  # Falls als .exe ausgeführt
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def create_labels(names, title, logo_path_1=None, logo_path_2=None, output_file_path='labels_with_logo.pdf'):
    # create canvas
    c = canvas.Canvas(output_file_path, pagesize=A4)
    width, height = A4

    # set Fonts parameter
    title_font = 'Helvetica-Oblique'
    title_font_size = 12
    names_font = 'Helvetica'
    names_font_size = 24
    line_spacing = 30
   
    # label size
    label_width = 9 * cm
    label_height = 5.5 * cm
    
    # logo 1 size (upper left)
    logo_1_width = 4 * cm
    logo_1_height = 4 * cm
    
    # logo 2 size (upper right)
    logo_2_width = 2 * cm
    logo_2_height = 2 * cm
    
    # border size
    border_left = 1 * cm
    border_top = 1 * cm
    
    # starting position with borders
    x = border_left
    y = height - label_height - border_top
    
    # fill in names from table
    for name in names:
        first_name, last_name = name
        
        # set stroke color to bright grey (RGB)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)

        # draw cut marks at the corners of the label
        # length of the cut marks
        cut_length = 0.5 * cm  

        # top-left corner (horizontal + vertical line)
        c.line(x, y + label_height, x + cut_length, y + label_height)
        c.line(x, y + label_height, x, y + label_height - cut_length)

        # top-right corner (horizontal + vertical line)
        c.line(x + label_width, y + label_height, x + label_width - cut_length, y + label_height)
        c.line(x + label_width, y + label_height, x + label_width, y + label_height - cut_length)

        # bottom-left corner (horizontal + vertical line)
        c.line(x, y, x + cut_length, y)
        c.line(x, y, x, y + cut_length)

        # bottom-right corner (horizontal + vertical line)
        c.line(x + label_width, y, x + label_width - cut_length, y)
        c.line(x + label_width, y, x + label_width, y + cut_length)

        # draw the first logo (if provided)
        if logo_path_1:
            logo_x1 = x + 0.2 * cm
            logo_y1 = y + label_height - logo_1_height - 0.2 * cm
            c.drawImage(logo_path_1, logo_x1, logo_y1, width=logo_1_width, height=logo_1_height, preserveAspectRatio=True, anchor='nw', mask='auto')

        # draw the second logo (if provided)
        if logo_path_2:
            logo_x2 = x + label_width - logo_2_width - 0.2 * cm
            logo_y2 = y + label_height - logo_2_height - 0.2 * cm
            c.drawImage(logo_path_2, logo_x2, logo_y2, width=logo_2_width, height=logo_2_height, preserveAspectRatio=True, anchor='nw', mask='auto')

        # adjust font size for the first name if it's too long (if provided)
        if pd.notna(title):
            while c.stringWidth(title, title_font, title_font_size) > label_width - 2 * cm and title_font_size > 8:
                title_font_size -= 1
        if pd.notna(first_name):
            while c.stringWidth(first_name, names_font, names_font_size) > label_width - 2 * cm and names_font_size > 8:
                names_font_size -= 1
        if pd.notna(last_name):
            while c.stringWidth(last_name, names_font, names_font_size) > label_width - 2 * cm and names_font_size > 8:
                names_font_size -= 1
        
        # calculate approximate text height only for provided elements
        title_height = pdfmetrics.getAscent(title_font) * (title_font_size / 1000) if pd.notna(title) else 0
        names_height = pdfmetrics.getAscent(names_font) * (names_font_size / 1000) if pd.notna(first_name) or pd.notna(last_name) else 0
        factor_pt_mm = 0.253

        # count active elements (title, first name, last name)
        active_elements = sum(pd.notna(x) for x in [title, first_name, last_name])
        total_text_height = (title_height + active_elements * names_height + (active_elements - 1) * line_spacing) * factor_pt_mm

        # draw the text dynamically based on available elements
        text_x = x + label_width / 2
        text_y = y + label_height / 2 + total_text_height / 2

        if pd.notna(title):
            c.setFont(title_font, title_font_size)
            c.drawCentredString(text_x, text_y, title)
            text_y -= line_spacing  # Move down for next element

        if pd.notna(first_name):
            c.setFont(names_font, names_font_size)
            c.drawCentredString(text_x, text_y, first_name)
            text_y -= line_spacing  # Move down for next element

        if pd.notna(last_name):
            c.setFont(names_font, names_font_size)
            c.drawCentredString(text_x, text_y, last_name)

        # move to the next label position
        x += label_width
        if x + label_width > width:
            x = border_left
            y -= label_height
            if y < border_top:
                c.showPage()
                y = height - label_height - border_top

    c.save()


def select_names_file(label):
    global names_file_path
    names_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if names_file_path:
        label.config(text=f"Selected file:\n{names_file_path}")
        df = pd.read_excel(names_file_path)
        # return column names
        return df.columns.tolist()
    return []


def on_select_names_file(label_select_file, label_title, input_title, first_name_combobox, last_name_combobox, first_name_label, last_name_label, button_select_logo_1, button_select_logo_2, button_create, button_open):
    # get columns from the select_file function
    columns = select_names_file(label_select_file)
    label_title.pack(pady=5)
    input_title.pack(pady=20)
    button_select_logo_1.pack(pady=5)
    button_select_logo_2.pack(pady=20)
    first_name_label.pack(pady=5)
    first_name_combobox.pack(pady=5)
    last_name_label.pack(pady=5)
    last_name_combobox.pack(pady=5)
    button_create.pack(pady=20)
    button_open.pack(pady=10)
    # update comboboxes with returned columns
    update_dropdowns(columns, first_name_combobox, last_name_combobox, first_name_label, last_name_label)  # Update comboboxes with returned columns


def update_dropdowns(columns, first_name_combobox, last_name_combobox, first_name_label, last_name_label):
    # fill dropdowns
    first_name_combobox['values'] = columns
    last_name_combobox['values'] = columns

    # set default values
    first_name_combobox.set(columns[0])
    last_name_combobox.set(columns[1])

    global selected_first_name, selected_last_name
    selected_first_name = first_name_combobox.get()
    selected_last_name = last_name_combobox.get()


def on_combobox_change(first_name_combobox, last_name_combobox):
    # check if both comboboxes have selections
    if first_name_combobox.get() and last_name_combobox.get():
        # set global vars with the selected values
        global selected_first_name, selected_last_name
        selected_first_name = first_name_combobox.get()
        selected_last_name = last_name_combobox.get()


def on_change_input_title(input_title):
    global title
    title = input_title.get()


def on_select_logo_1_file():
    global logo_1_file_path
    logo_1_file_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg;*.png")])


def on_select_logo_2_file():
    global logo_2_file_path
    logo_2_file_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg;*.png")])


def on_create_labels():
    # read the Excel file
    df = pd.read_excel(names_file_path, engine='openpyxl')

    # extract names as a list of tuples
    names = list(df[[selected_first_name, selected_last_name]].itertuples(index=False, name=None))

    # name of the PDF output file
    global output_file_path
    timestamp = datetime.now().strftime('%y%m%d_%H%M%S')
    output_file_path = f'{timestamp}_labels.pdf'

    # create the labels PDF with logo
    create_labels(
        names,
        title if title and len(title) > 0 else None,
        logo_1_file_path if logo_1_file_path else None,
        logo_2_file_path if logo_2_file_path else None,
        output_file_path
        )


def open_pdf(output_file_path):
    os.startfile(output_file_path)


# MAIN
def main():
    root = tk.Tk()
    root.title("Select Excel file")
    root.geometry("600x550")
    label_select_file = tk.Label(root, text="PLEASE SELECT AN EXCEL FILE WITH THE NAMES FOR THE LABELS")
    label_select_file.pack()

    # define the labels (for when the comboboxes are shown)
    first_name_label = tk.Label(root, text="Select First Name Column")
    last_name_label = tk.Label(root, text="Select Last Name Column")

    # define global var and the comboboxes
    global names_file_path, title, selected_first_name, selected_last_name, logo_1_file_path, logo_2_file_path, output_file_path
    title = None
    logo_1_file_path = None
    logo_2_file_path = None
    label_title = tk.Label(root, text="Title (optional)")
    input_title = tk.Entry(root)
    first_name_combobox = ttk.Combobox(root)
    last_name_combobox = ttk.Combobox(root)

    # initially hide the comboboxes and labels
    first_name_label.pack_forget()
    first_name_combobox.pack_forget()
    last_name_label.pack_forget()
    last_name_combobox.pack_forget()

    # bind items to events
    first_name_combobox.bind("<<ComboboxSelected>>", lambda event: on_combobox_change(first_name_combobox, last_name_combobox))
    last_name_combobox.bind("<<ComboboxSelected>>", lambda event: on_combobox_change(first_name_combobox, last_name_combobox))
    input_title.bind("<KeyRelease>", lambda event: on_change_input_title(input_title))

    # create buttons
    button_select_file = tk.Button(root, text="Select File", command=lambda: on_select_names_file(label_select_file, label_title, input_title, first_name_combobox, last_name_combobox, first_name_label, last_name_label, button_select_logo_1, button_select_logo_2, button_create, button_open))
    button_select_logo_1 = tk.Button(root, text="Select Logo 1", command=lambda: on_select_logo_1_file())
    button_select_logo_2 = tk.Button(root, text="Select Logo 2", command=lambda: on_select_logo_2_file())
    button_create = tk.Button(root, text="Create labels as PDF", command=lambda: on_create_labels())
    button_open = tk.Button(root, text="Open PDF", command=lambda: open_pdf(output_file_path))
    button_select_file.pack(pady=10)

    # initially hide items (that are not needed at this state)
    label_title.pack_forget()
    input_title.pack_forget()
    button_select_logo_1.pack_forget()
    button_select_logo_2.pack_forget()
    button_create.pack_forget()
    button_open.pack_forget()
   
    # start the tkinter event loop
    root.mainloop()


# starting main()
try:
    sys.exit(main())
except Exception as e:
    print(f"An error occurred: {e}")
    sys.exit(1)