#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import pandas as pd
from datetime import datetime
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):  # Falls als .exe ausgeführt
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def create_labels(names, title, logo_path_1=None, logo_path_2=None):

    timestamp = datetime.now().strftime('%y%m%d_%H%M%S')
    output_file_path = f'{timestamp}_labels.pdf'

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
    return output_file_path
