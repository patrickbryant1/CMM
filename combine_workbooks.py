#! /usr/bin/env python

import sys
import os
import pdb
import xlrd
import openpyxl
import glob
import argparse
from read_write import encode_ascii, write_to_sheet


#Arguments for argparse module:
parser = argparse.ArgumentParser(description = '''
This is a program that combines the first sheets in excel workbooks from a certain directory into one.

Two arguments: python code, directory with workbooks to combine

Note! Remember that one cannot write too many rows to an excel sheet (there
is a limit). 
'''
)
 
parser.add_argument('directory_path', nargs=1, type= str,
                  default=sys.stdin, help = 'path to directory containing excel files to be combined')

#Functions

def combine_sheets(directory_path):
    '''A functions that takes the path to a directory
    containing .xlsx workbooks and writes the information
    on the first sheets in them to a single sheet in a new
    workbook.
    Input: directory path
    Output: None
    '''

    row_w = 1 #To keep track of which row to write to
    row_r = 0 #To keep track of which row to rad from
    headers = False #Keep track of if headers have been written
    
    #Create an excel workbook and a sheet to write to
    workbook_w = openpyxl.Workbook()
    sheet_w = workbook_w.active #Get active sheet
    sheet_w.title = 'All_chr'   
    

    #Write over the first row of the first sheet in the first workbook 
    #(should be the same in all)

    for filename in glob.glob(os.path.join(directory_path, '*.xlsx')):

        
        #Open the first sheet in each workbook and write over all rows but the first
        workbook_r = xlrd.open_workbook(filename)
        sheet_r=workbook_r.sheet_by_index(0)
        if headers == False:
            write_to_sheet(row_r, sheet_w, sheet_r, row_w)
            headers = True
            
        for row_r in range(1, sheet_r.nrows): #Iterate through all rows but the first
            row_w +=1 #Write to the next row
            write_to_sheet(row_r, sheet_w, sheet_r, row_w)
                    

    workbook_w.save('summary.xlsx')

    return None
                
#Main program
args = parser.parse_args()

try:
    directory_path = args.directory_path[0]


except IOError:
    print 'No direcory path.'

try:
    combine_sheets(directory_path)
except IOError:
    print 'Could not combine sheets'



    
