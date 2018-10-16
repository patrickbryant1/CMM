#! /usr/bin/env python
# -*- coding: utf-8 -*-


import sys
import os
import pdb
import xlrd
import xlwt
import argparse
import openpyxl 


#Arguments for argparse module:
parser = argparse.ArgumentParser(description = '''A program that reads a file
containing haplotypes and aligns them at their correct genomic position
according to increasing window size.
Make sure all data from p-link is put in - otherwise the SNPs may
align incorrectly.''')
 
parser.add_argument('hap_file', nargs=1, type= str,
                  default=sys.stdin, help = 'path to .xls file with haplotypes to be opened')

parser.add_argument('instructions', nargs=1, type= str,
                  default=sys.stdin, help = '''path to instructions containing criteria to filter haplotypes on:
                    start, end, p_val and OR.''')

###########################################################################################
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

def hap_map(workbook_r, name, instructions):        
    '''A function that takes an xl-sheet as input and aligns
    its haplotypes according to increasing window size.
    Input: workbook_r, name
    Output: None
    '''

    #Get instructions
    out_start = instructions[0].split(' ')[1]
    out_end = instructions[1].split(' ')[1]
    p_threshold = instructions[2].split(' ')[1] #The p-value threshold
    OR_threshold = instructions[3].split(' ')[1]
    
    #Create an excel workbook and a sheet to write to
    workbook_w = openpyxl.Workbook()
    sheet_w = workbook_w.active #Get active sheet
    sheet_w.title = 'Visualized'
   
    for num in range(0, 1):#The first sheet, could have more
        sheet_r = workbook_r.sheet_by_index(num) #Open sheet_num, the one to read
        
	hap_lens = [] #List with haplotype lengths
        hap_lists = [] #List to store haplotype lists
        pos_to_rs = [] #Connect rs to pos


        for row_idx in range(1, sheet_r.nrows): #Start on row 2
            start = encode_ascii(sheet_r, row_idx, 3) #Get start position of haplotype (first SNP)
            end = encode_ascii(sheet_r, row_idx, 4) #Get end position of haplotype (second SNP)
            rs_1 = encode_ascii(sheet_r, row_idx, 5) #Get rs for start position
            rs_2 = encode_ascii(sheet_r, row_idx, 6) #Get rs for end position
            haplotype = encode_ascii(sheet_r, row_idx, 7) #Get haplotype
            F = encode_ascii(sheet_r, row_idx, 8) #Get frequqncy of haplotype in samples
            OR = encode_ascii(sheet_r, row_idx, 9) #Get odds ratio
            p_val = encode_ascii(sheet_r, row_idx, 11) #Get p-value
          
            start_rs = (start, rs_1) #Create tuple of start position and rs
            
            #Check output criteria are fulfilled'
            if p_val == 'NA':
                continue
            if float(out_start) <= float(start) and float(end) <= float(out_end): 
                if start_rs not in pos_to_rs:
                    pos_to_rs.append(start_rs) #Will only append unique
                if float(p_val) < float(p_threshold) and float(OR) > float(OR_threshold): #If the variant meets the significance and OR
                    hap_key = start + '/' + haplotype + '/' + rs_1 + '/' + rs_2 + '/' + p_val + '/' + F + '/' + OR 

                    (hap_lists, hap_lens) = hap_sort(hap_key, haplotype, hap_lens, hap_lists)

    #Write headers
    write_headers(sheet_w)
    #Write rs
    pos_to_rs = write_rs(pos_to_rs, sheet_w)
    #Write haps
    transpose(hap_lists, sheet_w, hap_lens, pos_to_rs)

    workbook_w.save('transposed_'+name[0])

    return None

def encode_ascii(xl_sheet, row_idx, col_idx):
    '''A function that encodes and retrieves excel cell values
    Input: row and column positions
    Output: cell_value
    '''
    cell_value = str(xl_sheet.cell(row_idx, col_idx).value).encode('ascii','ignore')
    return cell_value
    
def hap_sort(hap_key, haplotype, hap_lens, hap_lists):
    '''A function that creates lists and adds haplotypes to
    these lists according to the haplotype sizes
    Input = hap_key(str), haplotype(str), hap_lens(list), hap_lists(list)
    Output po= hap_lists(list), hap_lens(list)
    '''
    size = str(len(haplotype)) 

    
    #If the haplotype length has not been found
    if size not in hap_lens:
        hap_lens.append(size) #Be sure strs and ints match
        hap_lists.append([hap_key]) #Add hap_key   
        
    #If the haplotype length has been found
    else:
        for hap_list in hap_lists:
            if len(hap_list[0].split('/')[1]) == len(haplotype): #Haplotypes of equal size should be in the same list
                    hap_list.append(hap_key) #Make sure hap_keys are unique since there can be
                    break                         #many different haplotypes with equal size
            else:
                continue  # only executed if the inner loop did NOT break
            break  # only executed if the inner loop DID break    
    
    return hap_lists, hap_lens

def write_headers(sheet_w):
    '''A function that writes the headers to the
    workbook
    '''
    sheet_w.cell(row=5,column=1).value = 'Position' 
    sheet_w.cell(row=5,column=2).value = 'SNP1'

    write_vertical('P-value', 1, 2, sheet_w)
    write_vertical('OR', 2, 2, sheet_w)
    write_vertical('F', 3, 2, sheet_w)
    write_vertical('SNP2', 4, 2, sheet_w)

    return None

    
def take_first(item):
    '''Return first list element
    Input = item (list of lists)
    Output = item[0] (float)
    '''
    return float(item[0])
    
def write_rs(pos_to_rs, sheet_w):
    '''A function that writes positions and
    rs numbers to an excel sheet as well as
    indexing them
    '''
    
    #Sort on position
    pos_to_rs = sorted(pos_to_rs, key = take_first)
      
    for i in range(0,len(pos_to_rs)):
        #Write position
        item = pos_to_rs[i]
        cell = str(item[0]).encode('ascii','ignore')
        #sheet_w.write(i,0, cell) #Row i, column 0
        sheet_w.cell(row=i+6,column=1).value = cell #Row i, column 1
        #Write rs
        cell = str(item[1]).encode('ascii','ignore')
        sheet_w.cell(row=i+6,column=2).value = cell #Row i, column 2        
        pos_to_rs[i] = pos_to_rs[i] + (str(i+6),) #Add index to rs

    return pos_to_rs
        
def get_hap_len(item):
    '''A function that splits item on / and
    returns the haplotype length for sorting
    Input = item (list of lists)
    Output = length (int)
    '''
    
    length = len(item[0].split('/')[1])
    return length

def get_start(item):
    '''A function that splits item on / and
    returns the haplotype length for sorting
    Input = item (list of lists)
    Output = length (int)
    '''
    
    length = float((item.split('/')[0]))
    return length
     
    
def transpose(hap_lists, sheet_w, hap_lens, pos_to_rs):
    '''A function that writes data into a sheet in the excel workbook.
    Input: hap_dicts, sheet_w
    Output: None
    '''
    
    #Sort hap_lists on haplotype length
    hap_lists = sorted(hap_lists, key = get_hap_len)
    #Sort the contents of each hap_list on position
    for i in range(0, len(hap_lists)): #Remember to do a range when changing lists!
        hap_lists[i] = sorted(hap_lists[i], key = get_start)

    #When you print out the haps you want the ones of equal size in positional order
    col_idx = 2 #Keep track of column number
    for hap_list in hap_lists:
        for item in hap_list:
            item = item.split('/')
            haplotype = item[1]
            rs_1 = item[2]
            rs_2 = item[3]
            p_val = item[4]
            F = item[5]
            OR = item[6]

            for pr in pos_to_rs:
                if rs_1 == pr[1]: #Match on rs
                    row_idx = int(pr[2])
                    col_idx += 1 #Write to next column
                    write_haps(haplotype, rs_1, rs_2, p_val, F, OR, row_idx, col_idx, sheet_w)
                    break
    return None

def write_vertical(item, i, j, sheet_w):
    '''A function that writes the contents of item
    to th cell row i,col j in sheet_w.
    '''
    cell = item.encode('ascii','ignore')
    sheet_w.cell(row=i, column=j).value= cell
    sheet_w.cell(row=i, column=j).alignment = openpyxl.styles.Alignment(text_rotation = 90)

    return None

def write_haps(haplotype,rs_1, rs_2, p_val, F, OR, row_idx, col_idx, sheet_w):
    '''A function that writes the haplotype into an excel
    spread sheet.
    Input =
    Output = None
    '''

    #Write p-val
    write_vertical(p_val, 1, col_idx, sheet_w)
    #Write OR
    write_vertical(OR, 2, col_idx, sheet_w)
    #Write F
    write_vertical(F, 3, col_idx, sheet_w)
    #Write rs_2
    write_vertical(rs_2, 4, col_idx, sheet_w)
    #Write rs_1
    write_vertical(rs_1, 5, col_idx, sheet_w)

    for base in haplotype:
        cell = base.encode('ascii','ignore')
        sheet_w.cell(row=row_idx,column=col_idx).value = cell
        row_idx +=1

    return None
                
#Main program
     
args = parser.parse_args()

try:
    workbook_r =xlrd.open_workbook(args.hap_file[0])
    #The opened workbook is closed automatically after use
    
    #Checks if file is empty
    if os.stat(args.hap_file[0]).st_size == 0:
       print "Empty file."       
except IOError:
    print 'Cannot open', args.hap_file[0]
else:
    #Get criteria for output
    instructions = text_to_list(args.instructions[0])
    
    #Perform the alignment and ordering
    name = str(args.hap_file[0]).split('/')[-1:] #name is only the last part of path
                                            #name is a list here
    hap_map(workbook_r, name, instructions)


    
