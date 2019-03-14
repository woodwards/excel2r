# analyse workbook data
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)

# load functions
source("excel_crawler_functions.R")

# my separator (must be legal but not found in any wbname or wsname)
mysep <- "@@@@"

# identify numeric and percentages
x <- str_replace_all(data$value, "[,\\s]", "") # strip commas and spaces
# numbers -1.0E5, -2, 2,000.000
data$numeric <- as_numeric(x)
# percentages -5.0%, -5%
data$percent <- as_numeric(str_replace(str_match(x, "^-?[:digit:]+(\\.[:digit:]+)?%$")[,1], "%", ""))/100
# currency -$ 1,000.00
data$currency <- as_numeric(str_replace_all(str_match(x, "^-?\\$[:digit:]+(\\.[:digit:]+)?$")[,1], "\\$", ""))
# dates and times
# data$datetime <-
data <- data %>%
	mutate(
		values2=case_when(
			!is.na(numeric) ~ numeric,
			!is.na(percent) ~ percent,
			!is.na(currency) ~ currency,
			TRUE ~ NA_real_),
		formula2=case_when(
			is.na(formula) ~ "",
			TRUE ~ str_replace_all(formula, "[\\'\\$]", ""))
	)

# plots
ggplot() +
	geom_point(data=data, mapping=aes(x=col, y=row, size=str_length(formula2), colour=values2)) +
	scale_y_reverse() +
	facet_wrap( ~ str_c(wbname, wsname, sep="|"))

# https://stackoverflow.com/questions/27442991/how-to-get-a-hash-code-as-integer-in-r


# identify dependencies
# Excel worksheet names exclude \ / * ? : [ ] (and some other rules) maybe best to add manually
# Excel file paths exclude / * ? % | " < > (and some other rules) maybe best to add manually
# https://www.get-digital-help.com/2017/02/07/extract-cell-references-from-a-formula/

# define patterns
range <- "\\$?[A-Z]+\\$?[0-9]+(:\\$?[A-Z]+\\$?[0-9]+)?"
sheet <- "[a-zA-Z][a-zA-Z0-9\\s\\+\\-\\&\\_\\(\\)]*" # add any needed sheetname characters
book <- "\\[[a-zA-Z0-9][a-zA-Z0-9\\s\\+\\-\\&\\_\\.\\:\\\\]*\\]" # add any needed sheetname/bookname characters

# find all references
pattern <- paste0("('?((", book, ")?", sheet, ")'?!)?", range, "(?!\\()")
y <- str_extract_all(data$formula, pattern, simplify=TRUE) # matrix
y[is.na(y)] <- ""
data$dep <- do.call(paste, as_tibble(y))
# normalise
for (i in 1:nrow(y)){
	for (j in 1:ncol(y)){
		if (y[i,j]>""){
			y[i,j] <- str_replace_all(y[i,j], "'", "")
			y[i,j] <- str_replace_all(y[i,j], "^[^!]*$", paste0(data$wsname[i], "!", y[i,j]))
			y[i,j] <- str_replace_all(y[i,j], "^(?!\\[)", "[0]")
		}
	}
}
data$dep2 <- do.call(paste, as_tibble(y))
# translate
pattern <- paste0("(?<=\\])(", sheet, ")!")
for (i in 1:nrow(y)){
	for (j in 1:ncol(y)){
		if (y[i,j]>""){
			y[i,j] <- str_replace_all(y[i,j], "[$()]", "") # remove $
			y[i,j] <- str_replace_all(y[i,j], pattern, "_ws\\1_") # recode sheetname
			y[i,j] <- str_replace_all(y[i,j], "\\[([0-9]+)\\]", "wb\\1") # recode bookname
		}
	}
}
data$dep3 <- do.call(paste, as_tibble(y))



