#! /usr/bin/env python

'''
This is a program that combines the first sheets in excel workbooks from a certain directory into one.

Two arguments: python code, directory with workbooks to combine

Note! Remember that one cannot write too many rows to an excel sheet (there
is a limit). 
'''

import sys
import os
import pdb
import xlrd
import xlwt
import glob


#Functions

def combine_sheets(directory_path):
    '''A functions that takes the path to a directory
    containing .xlsx workbooks and writes the information
    on the first sheets in them to a single sheet in a new
    workbook.
    Input: directory path
    Output: None
    '''

    row_w = 0 #To keep track of which row to write to
    row_r = 0 #To keep track of which row to rad from
    headers = False #Keep track of if headers have been written
    
    #Create an excel workbook and a sheet to write to
    workbook_w = xlwt.Workbook()
    sheet_w = workbook_w.add_sheet('All_chr')
    
    

    #Write over the first row of the first sheet in the first workbook 
    #(should be the same in all)

    for filename in glob.glob(os.path.join(directory_path, '*.xlsx')):

        
        #Open the first sheet in each workbook and write over all rows but the first
        workbook_r = xlrd.open_workbook(filename)
        sheet_r=workbook_r.sheet_by_index(0)
        if headers == False:
            write_to_sheet(row_r, row_w, sheet_w, sheet_r)
            headers = True
            
        for row_r in range(1, sheet_r.nrows): #Iterate through all rows but the first
            row_w +=1 #Write to the next row
            write_to_sheet(row_r, row_w, sheet_w, sheet_r)
                    

    workbook_w.save('summary.xlsx')

    return None

def write_to_sheet(row_r, row_w, sheet_w, sheet_r):
    '''A function that writes data into a sheet in an excel workbook.
    Input: row_r, row_w, sheet_w, sheet_r
    Output: None
    '''

    col_n = 0 #To keep track of which column to write to
    
    #Iterate over all columns
    for col_idx in range(0, sheet_r.ncols):
        sheet_r_cell = sheet_r.cell(row_r, col_idx).value
	if type(sheet_r_cell) == float:
		sheet_r_cell = str(sheet_r_cell)
        sheet_r_cell = sheet_r_cell.encode('ascii','ignore')
        sheet_w.write(row_w, col_idx, sheet_r_cell)
        col_n += 1
        
      
    return None
                
#Main program

try:
    directory_path = sys.argv[1]

except IOError:
    print 'No direcory path.'

try:
    combine_sheets(directory_path)
except IOError:
    print 'Could not merge sheets'



    
