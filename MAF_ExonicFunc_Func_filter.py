#! /usr/bin/env python
# -*- coding: utf-8 -*-

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
import argparse
import openpyxl
from read_write import text_to_list, encode_ascii, write_to_sheet


#Arguments for argparse module:
parser = argparse.ArgumentParser(description = '''
This is a program that takes an xl workbook, reads it and uses its
sheets to perform filtering as specified by text files.

-Remember that you cannot write too many rows to an excel sheet,
there is a limitation!
''')
 
parser.add_argument('workbook_r', nargs=1, type= str,
                  default=sys.stdin, help = 'path to excel file with variants to be opened')

parser.add_argument('func_arg', nargs=1, type= str,
                  default=sys.stdin, help = 'path to excel file with funcs to be filered on')

parser.add_argument('exonic_func_arg', nargs=1, type= str,
                  default=sys.stdin, help = 'path to excel file with exonic funcs to be filered on')

parser.add_argument('ref_dbs', nargs=1, type= str,
                  default=sys.stdin, help = '''path to text file containing
                  reference databases and frequencies to filter on.''')

###########################################################################################

#Functions


def filter_sheet(workbook_r, name, func_arg, exonic_func_arg, ref_dbs):        
    '''A function that takes an xl-sheet
    as input and finds the shared variants as specified in zygosity positions
    Input: workbook_r, name, ref_dbs, zygosity_positions
    Output: None
    '''
    
    #Create an excel workbook and a sheet to write to
    workbook_w = openpyxl.Workbook()
    sheet_w = workbook_w.active #Get active sheet
    sheet_w.title = 'MAF_ExonicFunc_Func_filtered_'

    for num in range(0, 1):#workbook_r.nsheets):
        sheet_r = workbook_r.sheet_by_index(num) #Open sheet_num	
        row_idx = 0 #Row to transfer from
        row_n = 1 #To keep track of which row to write to
        #transfer first row
        write_to_sheet(row_idx, sheet_w, sheet_r, row_n)


        for row_idx in range(1, sheet_r.nrows):
            func = encode_ascii(sheet_r, row_idx, 4) #filter on Func.refGene (that is exonic, intronic, splicing etc)
            exonic_func = encode_ascii(sheet_r, row_idx, 5) #filter on ExonicFunc.refGene (that is frameshift_deletion/insertion, synonymous_SNV etc)
            
            if func not in func_arg and exonic_func not in exonic_func_arg:
                count_match = 0 #To keep track of how many db fulfill the criteria
                for ref_db in ref_dbs:
                    ref_db = ref_db.split(' ') #split each list item on space
                    MAF = encode_ascii(sheet_r, row_idx, int(ref_db[1]))#filter on reference databases MAFs
                    if MAF and MAF != '.': #Checks if MAF is empty or NA
                        MAF = float(MAF)
                        if(MAF >= float(ref_db[3]) or MAF <= float(ref_db[2])): #Filtering on MAF
                            count_match+=1
                    else:
                        count_match+=1 #If there is no MAF, it is important to keep
                if count_match == len(ref_dbs):
                    row_n += 1 #Write to next row
                    write_to_sheet(row_idx, sheet_w, sheet_r, row_n)                    
            

    workbook_w.save('MAF_ExonicFunc_Func_filtered'+name[0])

    return None

                
#Main program
args = parser.parse_args()

try:
    workbook_r =xlrd.open_workbook(args.workbook_r[0])
    #The opened workbook is closed automatically after use
    
    #Checks if file is empty
    if os.stat(args.workbook_r[0]).st_size == 0:
       print "Empty file."       
except IOError:
    print 'Cannot open', args.workbook_r[0]
else:


    #Arguments to filter on
    func_arg = text_to_list(args.func_arg[0])
    exonic_func_arg = text_to_list(args.exonic_func_arg[0])
    ref_dbs = text_to_list(args.ref_dbs[0])
    name = str(args.workbook_r[0]).split('/')[-1:] #name is only the last part of path
                                            #name is a list here
    filter_sheet(workbook_r, name, func_arg, exonic_func_arg, ref_dbs)


    
