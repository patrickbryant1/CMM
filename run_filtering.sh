#A script that filters xl files based on arguments given in external text files.
#The script automatically filters on MAF<=20% and MAF>=80%

for file in path_to_files
do
/home/pbryant/CMM/SNP_summary/mycode/jun_21 $file func exonic_func
done

echo This folder contains .xls files that have been filtered on the options specified in the text files here. > info.txt



