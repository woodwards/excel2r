# excel crawler using XLConnect

options(java.parameters="-Xmx1024m") # java memory
library(tidyverse)
library(XLConnect)

# potentially multiple wbs and ws
# wbname <- "Copy of NBO 2019 _Final_withGHG.xlsx"
wbname <- "Simplified forecaster.xlsx"

# functions
make_address <- function(row, col){
	paste0(ifelse(col<=26, LETTERS[col],paste0(LETTERS[col %/% 26],LETTERS[col %% 26])), row)
}

# http://altons.github.io/rstats/2015/02/13/quick-intro-to-xlconnect/
wb <- loadWorkbook(wbname, create=FALSE)
ws <- getSheets(wb)
value <- readWorksheet(wb, ws[1], header=FALSE, colTypes="character")
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
