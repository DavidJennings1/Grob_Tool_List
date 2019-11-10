'''Filter Grob Tools from master library and output as .pdf'''

# import os
# import re
# from collections import Counter
# import tkinter as tk
# from tkinter import ttk
from tkinter import filedialog  # noqa: F401
# from openpyxl import Workbook
from openpyxl import load_workbook
# from openpyxl.styles import Border, Side, Alignment, Font, NamedStyle


class Grob_Tools:

    def __init__(self):
        '''Gets CT number, description, holder and holder data
        from master tool list file.'''
        tl = 'King Machine Cutting Tool List.xlsx'  # noqa: E501
        wb = load_workbook(filename=tl)
        sh1 = wb.active
        temp = zip(sh1['A'], sh1['E'], sh1['C'], sh1['F'], sh1['J'], sh1['M'],
                   sh1['N'], sh1['P'], sh1['Q'], sh1['S'], sh1['Y'])
        print(list(temp))
        for t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11 in temp:
            if t2.value == 'MC16':
                print(t1.value, t3.value, t4.value, t5.value, t6.value,
                      t7.value, t8.value, t9.value, t10.value, t11.value)

        # for key in in_data:
        #     for cell1, cell2, cell3, cell4, cell5 in zip(sh1['A'], sh1['E'],
        #                                                  sh1['C'], sh1['J'],
        #                                                  sh1['Y']):
        #         if cell1.value == key and cell2.value == self.machine:
        #             t_data[key] = (cell3.value, cell4.value, cell5.value)
        # return t_data


try1 = Grob_Tools()
