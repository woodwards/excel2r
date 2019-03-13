# excel crawler using XLConnect
# Simon Woodward, DairyNZ 2019

options(java.parameters="-Xmx4g") # java memory
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk-11.0.2') # for 64-bit version
# options(java.parameters="-Xmx1024m") # java memory
# Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jre1.8.0_201') # for 32-bit version
library(rJava)
getOption("java.parameters")

# http://altons.github.io/rstats/2015/02/13/quick-intro-to-xlconnect/
# https://stackoverflow.com/questions/7963393/out-of-memory-error-java-when-using-r-and-xlconnect-package
# options(java.parameters="-Xmx1024m") # java memory
options(warn=-1)
library(tidyverse)
library(XLConnect)

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# workbooks to read
# wbnames <- c("Simplified forecaster.xlsx")
wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

#  create overarching data list
wb <- as.list(wbnames) # named list of workbooks
names(wb) <- wbnames

# add workbook
for(wbname in wbnames){
	print(paste("Adding workbook", wbname))
	stopifnot(str_detect(wbname, mysep)==FALSE)
	wb[[wbname]] <- list() # named list of worksheets
	wbo <- loadWorkbook(wbname, create=FALSE) # connection to workbook object
	wsnames <- getSheets(wbo)
	wb[[wbname]] <- as.list(wsnames) # named list of worksheets
	names(wb[[wbname]]) <- wsnames
	# i <- which(wsnames=="Chart1") # next one
	# wsnames <- wsnames[i:length(wsnames)]
	for (wsname in wsnames){
		print(paste("Adding worksheet", wsname))
		stopifnot(str_detect(wsname, mysep)==FALSE)
		value <- readWorksheet(wbo, wsname, header=FALSE, colTypes="character")
		formula <- value
		if (ncol(value)>0 && nrow(value)>0){
			for (col in 1:ncol(value)){
				for (row in 1:nrow(value)){
					if (!is.na(value[row,col])){
						# read formula
						# formula[row,col] <- "constant"
						try(
							formula[row,col] <- getCellFormula(wbo, wsname, row, col),
							silent=TRUE
						)
					}
				}
			}
		}
		fname <- paste0("temp/", wbname, mysep, wsname, ".rds")
		saveRDS(list(wbname=wbname, wsname=wsname, value=value, formula=formula),
				fname)
		xlcFreeMemory() # maybe this will help
	}
}
