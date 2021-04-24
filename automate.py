# -*- coding: utf-8 -*-
"""
Created on Wed Mar 31 17:59:20 2021

@author: Farhan Ashraf
"""
from docx import Document
import re
import os

import xlrd

from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Pt

path_of_excel_path = 'C:/Users/Farhan Ashraf/Desktop/holderID.xlsx'
excel_having_holder_id = xlrd.open_workbook(path_of_excel_path)
sheet = excel_having_holder_id.sheet_by_index(0)
holder_id = int(sheet.cell_value(1,0))

path_of_txt_file = 'C:/Users/Farhan Ashraf/Desktop/'
files = os.listdir(path_of_txt_file)

document = Document()

def change_orientation():
    current_section = document.sections[-1]
    new_width, new_height = current_section.page_height, current_section.page_width
    new_section = document.add_section(WD_SECTION.NEW_PAGE)
    new_section.orientation = WD_ORIENT.LANDSCAPE
    new_section.page_width = new_width
    new_section.page_height = new_height
    

    return new_section

for i in files:
    if i == 'operations' + str(holder_id) + '.txt':
        
        document.add_heading(i,0)
        txt_file_content = open('C:/Users/Farhan Ashraf/Desktop/' + i).read()
        word_file = document.add_paragraph(txt_file_content)
        
        font = document.styles['Normal'].font
        font.name = 'Courier New'
        font.size = Pt(8)
        print("Entered")
    
document.save('C:/Users/Farhan Ashraf/Desktop/' + i[:-4] + '.docx')
        

