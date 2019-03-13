# load workbook data from rds files
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)

# load functions
source("excel_crawler_functions.R")

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# workbooks to read
wbnames <- c("Simplified forecaster.xlsx")
# wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

#  initialise overarching dataframe
data <- setNames(vector("list", length(wbnames)), wbnames)
temp_files <- list.files(path="temp", pattern="*.rds")

# loop through workbooks
wbname <- wbnames[1]
for(wbname in wbnames){
	print(paste("Adding workbook", wbname))
	stopifnot(str_detect(wbname, mysep)==FALSE)
	# loop through workbookfiles=worksheets
	wsfiles <- str_match(temp_files, paste0(wbname, mysep, ".+rds")) # could fail if special characters?
	wsfiles <- wsfiles[!is.na(wsfiles)]
	wbdata <- setNames(vector("list", length(wsfiles)), wsfiles)
	wsfile <- wsfiles[1]
	for (wsfile in wsfiles){
		wsname <- str_match(wsfile, paste0("(?<=", wbname, mysep, ").+(?=.rds)"))[1]
		print(paste("Adding worksheet", wsname))
		stopifnot(str_detect(wsname, mysep)==FALSE)
		temp <- readRDS(paste0("temp/", wsfile))
		# loop through worksheet columns
		n <- ncol(temp$value)
		if (n>0){
			wsdata <- vector("list", n) # list of columns
			i <- 1
			for (i in 1:n){
				wsdata[[i]] <- tibble(
					wbname=temp$wbname,
				  	wsname=temp$wsname,
					code=codes[i],
					col=i,
					row=1:nrow(temp$value),
					value=temp$value[,i],
					formula=temp$formula[,i]
					) %>%
					filter(!is.na(value))
				# add left label
				# add top label
				# look for repeated formulas
				# other matrix analysis before it gets collapsed
			} # next col
			wbdata[[wsfile]] <- bind_rows(wsdata)
		}
	} # next wsfile
	data[[wbname]] <- bind_rows(wbdata)
} # next file/sheet
data <- bind_rows(data) %>%
	mutate(
		wbname=as.factor(wbname),
		wsname=as.factor(wsname),
		formula=str_replace(formula, "^\\+", "") # strip leading +
		)

# write_csv(select(data, -wbname), "data.csv")
