# excel crawler using XLConnect
# Simon Woodward, DairyNZ 2019

# http://altons.github.io/rstats/2015/02/13/quick-intro-to-xlconnect/
# https://stackoverflow.com/questions/7963393/out-of-memory-error-java-when-using-r-and-xlconnect-package
options(java.parameters="-Xmx1024m") # java memory
# options(java.parameters="-Xmx4g") # java memory causes a crash, maybe need to upgrade Java
options(warn=-1)
library(tidyverse)
library(XLConnect)

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# workbook to read (could be multiple)
wbname <- "Copy of NBO 2019 _Final_withGHG.xlsx"
# wbname <- "Simplified forecaster.xlsx"

#  create overarching data list
wb <- list() # named list of workbooks

# add workbook
print(paste("Adding workbook", wbname))
wb[[wbname]] <- list() # named list of worksheets
stopifnot(str_detect(wbname, mysep)==FALSE)
wbo <- loadWorkbook(wbname, create=FALSE) # connection to workbook object
wsnames <- getSheets(wbo)
names(wb[[wbname]]) <- wsnames # named list of worksheets
i <- which(wsnames=="Energy+1") # problem child!
wsnames <- wsnames[i:length(wsnames)]
for (wsname in wsnames){
	# named list of dataframes for value, formula, type etc for each worksheet
	print(paste("Adding worksheet", wsname))
	stopifnot(str_detect(wsname, mysep)==FALSE)
	value <- readWorksheet(wbo, wsname, header=FALSE, colTypes="character")
	formula <- value
	for (col in 1:ncol(value)){
		for (row in 1:nrow(value)){
			if (!is.na(value[row,col])){
				# read formula
				formula[row,col] <- "constant"
				try(
					formula[row,col] <- getCellFormula(wbo, wsname, row, col),
					silent=TRUE
				)
			}
		}
	}
	fname <- paste0("temp/", wbname, mysep, wsname, ".rds")
	saveRDS(list(wbname=wbname, wsname=wsname, value=value, formula=formula),
			fname)
	xlcFreeMemory() # maybe this will help
}

stop()

# functions
make_address <- function(row, col){
	paste0(ifelse(col<=26, LETTERS[col],paste0(LETTERS[col %/% 26],LETTERS[col %% 26])), row)
}

# dependencies (think about what I'm trying to do here)
role <- type
for (col in 1:ncol(value)){
	for (row in 1:nrow(value)){
		if (type[row,col] %in% c("numeric", "percentage")){
			if (formula[row,col]=="constant"){
				role[row,col] <- "input"
			} else {
				dependencies <- as.vector(str_match_all(formula[row,col], "\\$?[:upper:]{1,2}\\$?[:digit:]+")[[1]])
				if (length(dependencies)==0){
					role[row,col] <- "constantcalculation"
				} else {
					if (role[row,col]==type[row,col]){
						role[row,col] <- "calculation"
					}
					for (d in dependencies){
						daddress <- str_replace_all(d, "\\$", "")
						drow <- as.integer(str_match(d, "[:digit:]+")[[1]])
						dcol <- which(address[drow,]==daddress)
						if (!type[drow,dcol] %in% c("numeric", "percentage")){
							stop("reference to nonnumeric cell")
						}
						if (role[drow,dcol]=="calculation"){
							role[drow,dcol] <- "intermediate"
						}

					} # next d
				}
			}
		}
	} # next row
} # next col
