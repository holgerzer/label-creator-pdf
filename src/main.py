#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script Name: main.py
Description: A script to create labels on a PDF page.
Author: Holger Zernetsch
Date: 2025-02-17
Version: 0.5
License: GNU General Public License v3.0 or later (GPL-3.0+)
"""

from tkinter import Tk
from gui import LabelApp

def main():
    root = Tk()
    app = LabelApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
