#!/usr/bin/python2.7
# -*- coding: utf-8 -*-

import Tkinter as Tk
from tkFileDialog import askopenfilename
import ttk
from analyzers import ReviewTextAnalyzer
from functools import partial

__author__ = 'Jeff'

filetype = [('csv', '*.csv')]

root = Tk.Tk()
filename_var = Tk.StringVar(root, '')
category_var = Tk.StringVar(root, 'PERSONAL_CARE')


def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)


def open_file_dialog():
    filename = askopenfilename(filetypes = filetype)
    filename_var.set(filename)


def submit_input_file(button):
    try:
        button.config(state='disabled')
        analyzer = ReviewTextAnalyzer()
        #Init category
        analyzer.init_category(category_var.get())
        #Init input file path
        analyzer.set_input_file_path(filename_var.get())
        analyzer.analyze()
        analyzer.output()
        button.config(state='normal')
    except Exception:
        button.config(state='normal')

sub_panel0 = Tk.Frame(root, border=4)
sub_panel0.pack(side='top', anchor='w')
sub_panel1 = Tk.Frame(root, border=4)
sub_panel1.pack(side='top', anchor='w')
sub_panel2 = Tk.Frame(root, border=4)
sub_panel2.pack(side='top', anchor='w')
sub_panel3 = Tk.Frame(root, border=4)
sub_panel3.pack(side='top', anchor='w')


root.title('评论评分程序')
center_window(root, 600, 240)

#Sub panel0 components
Tk.Label(sub_panel0, text="Input File:  ").pack(side='left')
Tk.Button(sub_panel0, text='Browse', command=open_file_dialog).pack(side='left')

Tk.Label(sub_panel1, text='Category:  ').pack(side='left')
category_values = ['PERSONAL_CARE', 'BABY_CARE', 'HAIR_CARE', 'ORAL_CARE', 'FABRIC_CARE', 'FEMININE_CARE',
                   'APPLICANCES', 'PRESTIGE', 'SHAVE_ARE', 'SKIN_CARE']
category = ttk.Combobox(sub_panel1, textvariable=category_var, values=category_values).pack(side='left')

#Sub panel1 components
input_file_path_label = Tk.Label(sub_panel2, textvariable=filename_var).pack(side='left')

#Sub panel2 components
submit_button = Tk.Button(sub_panel3, text='Submit', command=lambda : submit_input_file(submit_button))
submit_button.pack(side='left')

root.mainloop()
