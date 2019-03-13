# load workbook data from rds files
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# workbooks to read
# wbnames <- c("Simplified forecaster.xlsx")
wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

#  create overarching data list
data <- setNames(vector("list", length(wbnames)), wbnames)

# add workbook
temp_files <- list.files(path="temp", pattern="*.rds")
wbname <- wbnames[1]
for(wbname in wbnames){
	print(paste("Adding workbook", wbname))
	stopifnot(str_detect(wbname, mysep)==FALSE)
	data[[wbname]] <- list() # initialise named list of worksheets
	wbfiles <- str_match(temp_files, paste0(wbname, mysep, ".+rds")) # could fail if special characters?
	wbfiles <- wbfiles[!is.na(wbfiles)]
	wbfile <- wbfiles[1]
	for (wbfile in wbfiles){
		wsname <- str_match(wbfile, paste0("(?<=", wbname, mysep, ").+(?=.rds)"))[1]
		print(paste("Adding worksheet", wsname))
		stopifnot(str_detect(wsname, mysep)==FALSE)
		data[[wbname]][[wsname]] <- readRDS(paste0("temp/", wbfile))
	} # next wbfile
} # next data

# this is so fast!!!!

