# excel crawler using XLConnect
# Simon Woodward, DairyNZ 2019

# https://stackoverflow.com/questions/7963393/out-of-memory-error-java-when-using-r-and-xlconnect-package
options(java.parameters="-Xmx1024m") # java memory
library(tidyverse)
library(XLConnect)

# functions
make_address <- function(row, col){
	paste0(ifelse(col<=26, LETTERS[col],paste0(LETTERS[col %/% 26],LETTERS[col %% 26])), row)
}

# wblist(wslist(ws(layer,layer,layer,...),ws,ws,...),wb,wb,wb,...)
# could be other attributes of workbooks (e.g. global names) and worksheets ...
# potentially multiple wbs and ws
# http://altons.github.io/rstats/2015/02/13/quick-intro-to-xlconnect/
wbname <- "Copy of NBO 2019 _Final_withGHG.xlsx"
# wbname <- "Simplified forecaster.xlsx"

#  create overarching data list
wb <- list() # named list of workbooks

# add workbook
print(paste("Adding workbook", wbname))
wb[[wbname]] <- list() # named list of worksheets
attributes(wb[[wbname]]) <- list("obj" = loadWorkbook(wbname, create=FALSE)) # connection to workbook object
wbo <- attributes(wb[[wbname]])$obj # convenience pointer to workbook connection
wsnames <- getSheets(wbo)
for (wsname in wsnames){
	# named list of dataframes for value, formula, type etc for each worksheet
	print(paste("Adding worksheet", wsname))
	wb[[wbname]][[wsname]] <- list("val" = readWorksheet(wbo, wsname, header=FALSE, colTypes="character"))
}

stop()

formula <- value
type <- value
address <- value
for (col in 1:ncol(value)){
	for (row in 1:nrow(value)){
		address[row,col] <- make_address(row,col)
		if (!is.na(value[row,col])){
			formula[row,col] <- "constant"
			try(
				formula[row,col] <- getCellFormula(wb, ws[1], row, col),
				silent=TRUE
			)
			type[row,col] <- case_when(
				str_detect(value[row,col], "%$") ~ "percentage",
				is.numeric(type.convert(value[row,col])) ~ "numeric",
				is.logical(type.convert(value[row,col])) ~ "logical",
				TRUE ~ "string")
			if (type[row,col]=="percentage"){
				value[row,col] <- as.numeric(str_replace(value[row,col], "%$", ""))/100
			}
		}
	}
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
