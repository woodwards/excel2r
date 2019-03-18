# read excel workbooks and save data to rds
# Simon Woodward, DairyNZ 2019

# http://altons.github.io/rstats/2015/02/13/quick-intro-to-xlconnect/
# https://stackoverflow.com/questions/7963393/out-of-memory-error-java-when-using-r-and-xlconnect-package
options(warn=-1) # ignore warnings to save memory
if (.Machine$sizeof.pointer==4){
	options(java.parameters="-Xmx1024m") # increase java memory
	Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jre1.8.0_201') # for 32-bit version
} else if (.Machine$sizeof.pointer==8){
	options(java.parameters="-Xmx4g") # increase java memory
	Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk-11.0.2') # for 64-bit version
}

# load libraries
library(tidyverse)
library(XLConnect)

# workbooks to read
# wbnames <- c("Simplified forecaster.xlsx")
# wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

# make column codes vector
xlcol <- 1:10000 # or as wide as necessary, max is 16384
xlcode <- if_else(xlcol<=26L, LETTERS[xlcol], paste0(LETTERS[pmax(1, (xlcol-1L) %/% 26L)], LETTERS[(xlcol-1L) %% 26L + 1L]))
names(xlcol) <- xlcode

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# add workbook
wbname <- wbnames[1] # for testing
for(wbname in wbnames){
	print(paste("Adding workbook", wbname))
	stopifnot(str_detect(wbname, mysep)==FALSE)
	wbo <- loadWorkbook(wbname, create=FALSE) # connection to workbook object
	wsnames <- getSheets(wbo)
	# use this to resume after crash
	# wsnames <- wsnames[which(wsnames=="Chart1"):length(wsnames)]
	wsname <- wsnames[2] # for testing
	for (wsname in wsnames){
		print(paste("Adding worksheet", wsname))
		stopifnot(str_detect(wsname, mysep)==FALSE)
		value <- readWorksheet(object=wbo,
							   sheet=wsname,
							   header=FALSE,
							   colTypes="character",
							   startRow=1,
							   startCol=1,
							   useCachedValues=TRUE, # avoids external reference errors
							   autofitRow=FALSE,
							   autofitCol=FALSE
		)
		if (ncol(value)>0 && nrow(value)>0){
			names(value) <- codes[1:ncol(value)]
			formula <- value
			for (col in 1:ncol(value)){
				formula[,col] <- NA_character_
				for (row in 1:nrow(value)){
					if (!is.na(value[row,col])){
						# read formula
						try(
							formula[row,col] <- getCellFormula(wbo, wsname, row, col),
							silent=TRUE
						)
					}
				}
			}
		} else {
			formula <- value
		}
		wbnum <- which(wbnames==wbname)
		wsnum <- which(wsnames==wsname)
		fname <- paste0("temp/", wbname, mysep, wsname, ".rds")
		saveRDS(list(wbname=wbname, wbnum=wbnum, wsname=wsname, wsnum=wsnum,
					 value=value, formula=formula),
				fname)
		xlcFreeMemory() # maybe this will help with memory
	}
}
