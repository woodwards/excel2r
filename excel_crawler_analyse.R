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

# identify dependencies
# Excel worksheet names exclude \ / * ? : [ ] (and some other rules) maybe best to add manually
# Excel file paths exclude / * ? % | " < > (and some other rules) maybe best to add manually
# https://www.get-digital-help.com/2017/02/07/extract-cell-references-from-a-formula/

# simple range
range <- "\\$?[A-Z]+\\$?[0-9]+(:\\$?[A-Z]+\\$?[0-9]+)?"
pattern <- range
y <- str_extract_all(data$formula, pattern, simplify=TRUE) # matrix
y[is.na(y)] <- ""
data$dep1 <- do.call(paste, as_tibble(y))
# sheet!range
sheet <- "[a-zA-Z][a-zA-Z0-9\\s\\+\\-\\&\\_\\(\\)]*" # add any needed sheetname characters
pattern <- paste0("'?", sheet, "'?!", range)
y <- str_extract_all(data$formula, pattern, simplify=TRUE) # matrix
y[is.na(y)] <- ""
data$dep2 <- do.call(paste, as_tibble(y))
# [path]sheet!range
book <- "\\[[a-zA-Z0-9][a-zA-Z0-9\\s\\+\\-\\&\\_\\.\\:\\\\]*\\]" # add any needed sheetname/bookname characters
pattern <- paste0("('?((", book, ")?", sheet, ")'?!)?", range)
y <- str_extract_all(data$formula, pattern, simplify=TRUE) # matrix
y[is.na(y)] <- ""
data$dep3 <- do.call(paste, as_tibble(y))

