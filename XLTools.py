# -*- coding: utf-8 -*-
"""
 Module to handle the data transfert with Microsoft Excel©
"""
___author___   = 'Cédric Montero'
___contact___  = 'cedric.montero@esrf.fr'
___copyright__ = '2012, European Synchrotron Radiation Facility'
___version___  = '0'

""" External modules (preliminary installation could be require) """
# Module to write Microsoft Excel© files
import xlwt
# Module to read Microsort Excel©  files 
import xlrd

def XL_export(filename,data_array,title=[]):
    """
    Export data to Microsoft Excel© file
	!!! : Be careful it will overright the existing file whout preventing
	#TODO : Check if the file exist and ask user want to do (overright, rename, cancel)
    @param filename : name of the file with extension .xls
    @type filename: string
    @param data_array : list of array to save in the columns of the sheet
    @type data_array : list of numpy 1D array
    @param title : headers of first line of cells in the sheet
    @type title : list of string
    """
    wb = xlwt.Workbook() # Create a Excel workbook
    ws = wb.add_sheet("Data") # Create a sheet
    line = 0
    col = 0
    #Create XL line 1 with titles
    for texte in title:
        ws.write(0,col,texte)
        col = col+1
    #Find datetime columns :
    style = xlwt.XFStyle()#Style des datetime objects dans Excel
    style.num_format_str = 'M/D/YY h:mm'
    # Other options of format : D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
    #if ... == 
    #Write the data
    line = 1
    col = 0
    for vector in data_array:
        for val in vector:
            ws.write(line,col,val)
            line=line+1
        col = col+1
        line = 1
    #Save the workbook
    wb.save(filename)

def XL_create_empty(filename,sheet1_name):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet1_name)
    wb.save(filename)

def XL_export_in_existing(filename,data_array,col,title=[]):
    rb = xlrd.open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    r_sheet.write(2,2,'glr')
    wb.save(filename)
