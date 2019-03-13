# analyse workbook data
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)

# load functions
source("excel_crawler_functions.R")

# loop through sheets
wbname <- names(data)[1]
for (wbname in names(data)){
	print(paste("Analysing workbook", wbname))
	wsname <- names(data[[wbname]])[1]
	wsname <- "Inputs Industry "
	for (wsname in names(data[[wbname]])){
		print(paste("Analysing worksheet", wsname))

		# prepare to analyse ws
		value <- data[[wbname]][[wsname]]$value
		form <- data[[wbname]][[wsname]]$formula

		# dplyr is faster than loop
		type <- value %>%
			mutate_all(what_type) %>%
			mutate_all(as.factor)
		inputs <- unlist(find_inputs(type, form))

	    # report results
	    if (length(inputs)>0){
	    	print(paste("Inputs", paste(inputs, collapse=" ")))
	    }

	    # store results
	    data[[wbname]][[wsname]][["inputs"]] <- inputs
	    data[[wbname]][[wsname]]$type <- type

	} # next ws
} # next wb


