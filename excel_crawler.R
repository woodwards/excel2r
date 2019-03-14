# analyse workbook data
# Simon Woodward, DairyNZ 2019

# workbooks to read
wbnames <- c("Simplified forecaster.xlsx")
# wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

# read excel and save to rds
source("excel_crawler_read.R")

# load rds and construct dataframe
source("excel_crawler_load.R")

# simple parse of values and formulae and create visualisation
source("excel_crawler_visualise.R")
