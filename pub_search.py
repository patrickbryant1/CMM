#! /usr/bin/env python
from Bio import Entrez
import pdb
import sys
import argparse


#Arguments for argparse module:
parser = argparse.ArgumentParser(description = '''A program that reads a list of genes
and searches pubmed for information related to each gene. The found information
is then written to stdout''')
 
parser.add_argument('infile', nargs=1, type= str,
                  default=sys.stdin, help = '''path to file with a list of newline
                    separated genes to be opened''')

# *Always* tell NCBI who you are
Entrez.email = "patrick.bryant@ki.se"

def text_to_list(file_name):
    '''A function that reads a text file and creates a list of its rows.
    Input: text_file (txt)
    Output: text_list (list)
    '''

    text_list = [] #empty list to store options in
    with open(file_name, 'r') as infile:
        for line in infile:
            line = line.rstrip('\n')
            text_list.append(line)
            
    return text_list
    

#Get info on db parameters
#data = Entrez.read(Entrez.einfo(db="pubmed"))
#for field in data["DbInfo"]["FieldList"] :
#    print "%(Name)s, %(FullName)s, %(Description)s" % field



def gene_search(gene_list, dbs):
    '''A function that searches the databases in dbs
    for the gene names in gene_list.
    input = gene_list(list),dbs(dict)
    output = none
    '''
    
    
    for single_term in gene_list:
        for key in dbs:
            #Fetch ids
	    handle = Entrez.esearch(db=key,term = single_term)
            record = Entrez.read(handle)
            ids = record['IdList']
            
            #Print the found info
            print 'Search term:', single_term
            print 'Database', key
            print 'Number of search results:', len(ids)
            n_disp = dbs[key] #Reset n_disp
            if len(ids)<dbs[key]: #If there are less results than wanted
            	n_disp = len(ids)

            print 'Number of search results displayed:', abs(n_disp)
            handle.close() #Close handle before creating a new one

            #Fetch article TIABs
            for article_id in ids[0:dbs[key]]:
                handle = Entrez.efetch(db=key, id=article_id, rettype="TIAB", retmode="text") #TIAB = Title/Abstract, Free text associated with Abstract/Title
                record = handle.read()

                print record

                handle.close()
    	print '*'*80, '\n','*'*80 #Separator

dbs = {'pubmed':3} #Dict with db to search and number of results to display.




#Main program
#Read gene list with argparse

#gene_list = ['MSH3']
args = parser.parse_args()
gene_list = text_to_list(args.infile[0])
gene_search(gene_list, dbs)#Search
