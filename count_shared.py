#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 20 13:54:41 2018

@author: patbry
"""

import pdb
'''This is a program that takes an xl workbook, reads it and uses its
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

#Arguments for argparse module:
parser = argparse.ArgumentParser(description = '''
This is a program that takes an xl workbook, reads it and uses its
sheets to perform filtering as specified by text files.

-Remember that you cannot write too many rows to an excel sheet,
there is a limitation!
''')
 
parser.add_argument('workbook_r', nargs=1, type= str,
                  default=sys.stdin, help = 'path to excel file with variants to be opened')


parser.add_argument('ref_dbs', nargs=1, type= str,
                  default=sys.stdin, help = '''path to text file containing column positions
                  for ref_db (not MAX_REF_MAF) to filter on.''')

parser.add_argument('zygosity_start_end', nargs=1, type= str,
                  default=sys.stdin, help = '''path to text file containing
                  zygosity start and end positions to count presence in.''')


from read_write import text_to_list, encode_ascii, write_to_sheet

def filter_sheet(workbook_r, name, ref_dbs, zygosity_start_end):        
    '''A function that takes an xl-sheet
    as input and finds the shared variants as specified in zygosity positions
    Input: workbook_r, name, ref_dbs, zygosity_start_end
    Output: None
    '''
    
    shared_genes = [] #Keep track of genes matched
    
    #Create an excel workbook and a sheet to write to
    workbook_w = openpyxl.Workbook()
    sheet_w = workbook_w.active #Get active sheet
    sheet_w.title = 'shared_count'


    for num in range(0, 1):#workbook_r.nsheets):
        sheet_r = workbook_r.sheet_by_index(num) #Open sheet_num
	
        row_idx = 0 #Row to transfer from
        row_n = 1 #To keep track of which row to write to
        #transfer first row
        write_to_sheet(row_idx, sheet_w, sheet_r, row_n)


        for row_idx in range(1, sheet_r.nrows):
            above_t = False

            ref_db = ref_dbs[0].split(' ') #split each list item on space
            MAF = encode_ascii(sheet_r, row_idx, int(ref_db[1]))#filter on reference databases MAFs
            if MAF and MAF != '.': #Checks if MAF is empty or NA
                MAF = float(MAF)
                if MAF > float(ref_db[3]): #Filtering on MAF above threshold
                        above_t = True
            
            shared_genes = count_shared(zygosity_start_end, sheet_r, row_idx, above_t, shared_genes)
           
    print shared_genes

    workbook_w.save('shared_'+name[0])

    return None

def count_shared(zygosity_start_end, xl_sheet, row_idx, above_t, shared_genes):
    '''Finds the variants shared between the samples as specified by zygosity_positions
        Input = zygosity_positions, xl_sheet, row_idx, above_t
        Output = shared (bool)
    '''
    

    #Get positions for zygosities
    start = int(zygosity_start_end[0].split(' ')[0])
    end =  int(zygosity_start_end[0].split(' ')[1])
    
    #Keep track of matches
    wt = 0 #Keep track of wt
    hom = 0 #Keep track of hom
    het = 0 #Keep track of het
    other = 0 #Keeep track of other
    
    #Search samples
    for pos in range(start-1,end):
        zyg = encode_ascii(xl_sheet, row_idx, pos) #zygosity to match
        if zyg == "wt" or not zyg: #If the zyg is empty, it is wt
            wt +=1
        if zyg == "hom":
            hom += 1
        if zyg == "het":
            het +=1
        if zyg == 'oth':
            other +=1
            
    #Count        
    if above_t == True:
        shared_count = het + wt + other
    else:
        shared_count = het + hom + other
    
    #Assess gene
    found = False
    gene = encode_ascii(xl_sheet, row_idx, 7)
    for i in range (0, len(shared_genes)):
        old = shared_genes[i].split('/')
        if gene == old[0]:
            shared_genes[i] = gene + '/' + str(int(old[1])+ shared_count)
            found = True
            break
    
    if found == False:
        shared_genes.append(gene + '/' + str(shared_count))
        
    return(shared_genes)
    
    
#Main program
args = parser.parse_args()

try:
    workbook_r = xlrd.open_workbook(args.workbook_r[0])
    #The opened workbook is closed automatically after use
    
    #Checks if file is empty
    if os.stat(args.workbook_r[0]).st_size == 0:
       print "Empty file."       
except IOError:
    print 'Cannot open', args.workbook_r[0]
else:


    #Arguments to filter on
    ref_dbs = text_to_list(args.ref_dbs[0])
    zygosity_start_end = text_to_list(args.zygosity_start_end[0])
    
    #Perform the filtering
    name = str(args.workbook_r[0]).split('/')[-1:] #name is only the last part of path
                                            #name is a list here
    filter_sheet(workbook_r, name, ref_dbs, zygosity_start_end)
