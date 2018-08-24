#A script that filters xl files based on arguments given in external text files.

for file in ~/CMM/PMS2_mutations/Summary/*
do
~/CMM/code/filtering/20180824/xlrd_filtering.py $file func exonic_func ref_dbs zygosity_positions
done

echo This folder contains .xls files that have been filtered on the options specified in the text files here. > info.txt



