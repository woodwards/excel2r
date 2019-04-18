# load workbook data from rds files
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)

# workbooks to read
# wbnames <- c("Simplified forecaster.xlsx")
# wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# make column codes vector
xlcol <- 1:10000 # or as wide as necessary, max is 16384
xlcode <- if_else(xlcol<=26L, LETTERS[xlcol], paste0(LETTERS[pmax(1, (xlcol-1L) %/% 26L)], LETTERS[(xlcol-1L) %% 26L + 1L]))
names(xlcol) <- xlcode

#  initialise overarching dataframe
data <- setNames(vector("list", length(wbnames)), wbnames)
temp_files <- list.files(path="temp", pattern="*.rds")

# loop through workbooks
wbname <- wbnames[1] # for testing
for(wbname in wbnames){
	print(paste("Adding workbook", wbname))
	stopifnot(str_detect(wbname, mysep)==FALSE)
	# loop through workbookfiles=worksheets
	wsfiles <- str_match(temp_files, str_c("\\Q", wbname, mysep, "\\E.+rds")) # \\Q and \\E bracket fixed portion of pattern
	wsfiles <- wsfiles[!is.na(wsfiles)]
	wbdata <- setNames(vector("list", length(wsfiles)), wsfiles)
	wsfile <- wsfiles[1]
	for (wsfile in wsfiles){
		wsname <- str_match(wsfile, str_c("(?<=\\Q", wbname, mysep, "\\E).+(?=.rds)"))[1]
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
					wbnum=temp$wbnum,
					wsname=temp$wsname,
					wsnum=temp$wsnum,
					code=xlcode[i],
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
	arrange(wbnum, wsnum, row) %>%
	mutate(
		wbname=factor(wbname, levels=unique(wbname)), # retain ordering
		wsname=factor(wsname, levels=unique(wsname)) # retain ordering
		)
rm(list=c("wsdata", "wbdata", "temp"))

# write_csv(select(data, -wbname), "data.csv")
