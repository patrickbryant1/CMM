#! /usr/bin/env python

'''
This is a program that takes an xl workbook, reads it and uses its
sheets to perform filtering as specified by the text file.

-Remember that you cannot write too many rows to an excel sheet,
there is a limitation!
'''

import sys
import os
import pdb
import xlrd
import xlwt

#Functions

def text_to_list(file_name):
    '''A function that reads a text file and creates a list of its rows.
    Input: text_file (txt)
    Output: filter_options (list)
    '''

    filter_options = [] #empty list to store options in
    with open(file_name, 'r') as infile:
        for line in infile:
            line = line.rstrip('\n')
            filter_options.append(line)
            
    return filter_options


def filter_sheet(workbook_r, name, func_arg, exonic_func_arg, ref_dbs):
    '''A function that takes an xl-sheet
    as input and filters it according to the users specifications
    on Func.refGene (=func), ExonicFunc.refGene (=exonic_func),
    and different reference databases as specified by ref_dbs.
    Input: workbook_r, name, func_arg, exonic_func_arg, ref_dbs
    Output: None
    '''
    
    #Create an excel workbook and a sheet to write to
    workbook_w = xlwt.Workbook()

    for num in range(0, workbook_r.nsheets):
        xl_sheet = workbook_r.sheet_by_index(num) #Open sheet_num
        sheet_w = workbook_w.add_sheet('Sheet_'+str(num+1)) #Create sheet_num into workbook_w

        row_idx = 0 #Row to transfer from
        row_n = 0 #To keep track of which row to write to
        #transfer first row
        write_to_sheet(row_idx, sheet_w, xl_sheet, row_n)


        for row_idx in range(1, xl_sheet.nrows):
            func = xl_sheet.cell(row_idx, 4).value.encode('ascii','ignore') #filter on Func.refGene (that is exonic, intronic, splicing etc)
            exonic_func = xl_sheet.cell(row_idx, 5).value.encode('ascii','ignore') #filter on ExonicFunc.refGene (that is frameshift_deletion/insertion, synonymous_SNV etc)
            if func not in func_arg and exonic_func not in exonic_func_arg:
                for ref_db in ref_dbs:
                    ref_db = ref_db.split(' ') #split each list item on space
                    MAF = xl_sheet.cell(row_idx, int(ref_db[1])).value.encode('ascii','ignore') #filter on reference databases MAFs
                    if MAF and MAF != '.': #Checks if MAF is empty or NA
                        MAF = float(MAF)
                        if(MAF >= float(ref_db[3]) or MAF <= float(ref_db[2])): #Filtering on MAF
                            row_n += 1
                            write_to_sheet(row_idx, sheet_w, xl_sheet, row_n)
                    else:
                        row_n += 1
                        write_to_sheet(row_idx, sheet_w, xl_sheet, row_n)                        
                    
            else:
                continue
    workbook_w.save('filtered_'+name[0])

    return None

def write_to_sheet(row_idx, sheet_w, xl_sheet, row_n):
    '''A function that writes data into a sheet in the excel workbook.
    Input: row_idx, sheet_w, xl_sheet
    Output: None
    '''

    #Iterate over all columns
    for col_idx in range(0, xl_sheet.ncols):
        cell = str(xl_sheet.cell(row_idx, col_idx).value)
        cell = cell.encode('ascii','ignore')
        sheet_w.write(row_n, col_idx, cell)

    return None
                
#Main program

try:
    workbook_r = xlrd.open_workbook(sys.argv[1])
    #The opened workbook is closed automatically after use
    
    #Checks if file is empty
    if os.stat(sys.argv[1]).st_size == 0:
       print "Empty file."       
except IOError:
    print 'Cannot open', sys.argv[1]
else:


    #Arguments to filter on

    func_arg = text_to_list(sys.argv[2])
    exonic_func_arg = text_to_list(sys.argv[3])
    ref_dbs = text_to_list(sys.argv[4])
    print ref_dbs

    #Perform the filtering

    name = str(sys.argv[1]).split('/')[-1:] #name is only the last part of path
                                            #name is a list here
    filter_sheet(workbook_r, name, func_arg, exonic_func_arg, ref_dbs)


    
