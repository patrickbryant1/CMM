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


 
def zygosity_match(zygosity_positions, xl_sheet, row_idx):
    '''A function that takes a list of zygosity positions
    and specified zygosity combinations and matches them.
    Input: list of zygosity  positions and combinations
    Output: riht_zyg = True/False, match or no_match
    '''
    column_pos = [] #List to store colummn positions
    right_zyg = False    
    
   
    for item in zygosity_positions:
        item = item.split(' ') #split on space
        column_pos.append(item[1])  #get column position

    match_count = 0 #To keep track of how many matches have been made
    x = len(zygosity_positions[0].split(' ')) #To keep track of last column
    #Match zygosities
    for i in range(2, x): #Go through all zygosity combinations
        j = 0 #go through all items
        match_count = 0
        for pos in column_pos:
            match = encode_ascii(xl_sheet, row_idx, int(pos)) #zygosity to match
            item = zygosity_positions[j].split(' ')
            zygosity = item[i]
            if match == zygosity:
                match_count +=1
            j+=1
        if match_count == len(column_pos): #Check if all columns have been matched
            right_zyg = True
            break

    return right_zyg

def encode_ascii(xl_sheet, row_idx, col_idx):
    '''A function that encodes and retrieves excel cell values
    Input: row and column positions
    Output: cell_value
    '''
    cell_value = str(xl_sheet.cell(row_idx, col_idx).value).encode('ascii','ignore')
    return cell_value
    

def filter_sheet(workbook_r, name, func_arg, exonic_func_arg, ref_dbs, zygosity_positions):        
    '''A function that takes an xl-sheet
    as input and filters it according to the users specifications
    on Func.refGene (=func), ExonicFunc.refGene (=exonic_func),
    and different reference databases as specified by ref_dbs.
    Input: workbook_r, name, func_arg, exonic_func_arg, ref_dbs
    Output: None
    '''
    
    #Create an excel workbook and a sheet to write to
    workbook_w = xlwt.Workbook()

    for num in range(0, 1):#workbook_r.nsheets):
        xl_sheet = workbook_r.sheet_by_index(num) #Open sheet_num
        sheet_w = workbook_w.add_sheet('Sheet_'+str(num+1)) #Create sheet_num into workbook_w

        row_idx = 0 #Row to transfer from
        row_n = 0 #To keep track of which row to write to
        #transfer first row
        write_to_sheet(row_idx, sheet_w, xl_sheet, row_n)


        for row_idx in range(1, xl_sheet.nrows):
            func = encode_ascii(xl_sheet, row_idx, 4) #filter on Func.refGene (that is exonic, intronic, splicing etc)
            exonic_func = encode_ascii(xl_sheet, row_idx, 5) #filter on ExonicFunc.refGene (that is frameshift_deletion/insertion, synonymous_SNV etc)
            right_zyg = zygosity_match(zygosity_positions, xl_sheet, row_idx) #Compare zygosities between individuals
            if func not in func_arg and exonic_func not in exonic_func_arg and right_zyg == True:
                count_match = 0 #To keep track of how many db fulfill the criteria
                for ref_db in ref_dbs:
                    ref_db = ref_db.split(' ') #split each list item on space
                    MAF = encode_ascii(xl_sheet, row_idx, int(ref_db[1]))#filter on reference databases MAFs
                    if MAF and MAF != '.': #Checks if MAF is empty or NA
                        MAF = float(MAF)
                        if(MAF >= float(ref_db[3]) or MAF <= float(ref_db[2])): #Filtering on MAF
                            count_match+=1
                    else:
                        count_match+=1 #If there is no MAF, it is important to keep
                if count_match == len(ref_dbs):
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
    zygosity_positions = text_to_list(sys.argv[5])
    
    #Perform the filtering

    name = str(sys.argv[1]).split('/')[-1:] #name is only the last part of path
                                            #name is a list here
    filter_sheet(workbook_r, name, func_arg, exonic_func_arg, ref_dbs, zygosity_positions)


    
