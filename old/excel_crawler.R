# analyse workbook data
# Simon Woodward, DairyNZ 2019

# workbooks to read
wbnames <- c("Simplified forecaster.xlsx")
# wbnames <- c("FD1004 Data For Modelling.xlsx")
# wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")
# wbnames <- c("Example MPI Model - Southern View 25 March.xlsx")

# read excel and save to rds
# source("excel_crawler_read.R")

# load rds and construct dataframe
source("excel_crawler_load.R")

# simple parse of values and formulae and create visualisation
source("excel_crawler_visualise.R")

# find terminal regions and work backwards to inputs!
# there is no unique sorting order
# attempt to name regions intelligently
# if regions reference themselves we need to split them
# sometimes a loop or iteration may be necessary
