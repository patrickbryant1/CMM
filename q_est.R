#!/usr/bin/env Rscript

#Import library
library(qvalue)
# import data
p <- read.delim("/home/patrick/CMM/results/BRC_haplotypes/20181010_test/histo/p_vals.txt", sep = '\n')

p <- as.numeric(as.character(unlist(p)))
hist(p)
# get q-value object
lambda = seq(0, 0.15, 0.01) #Look at the p-value distribution to set lambda max
qobj <- qvalue(p, lambda = lambda)
plot(qobj)
hist(qobj)
hist(qobj[["qvalues"]])
#write(q, file = "data.txt")

# options available
#qobj <- qvalue(p, lambda=0.5, pfdr=TRUE)
#qobj <- qvalue(p, fdr.level=0.05, pi0.method="bootstrap", adj=1.2)



