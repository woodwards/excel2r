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
data$numeric <- as.numeric(x)
# percentages -5.0%, -5%
data$percent <- as.numeric(str_replace(str_match(x, "^-?[:digit:]+(\\.[:digit:]+)?%$")[,1], "%", ""))/100
# currency -$ 1,000.00
data$currency <- as.numeric(str_replace_all(str_match(x, "^-?\\$[:digit:]+(\\.[:digit:]+)?$")[,1], "\\$", ""))
# dates and times
# data$datetime <-

# identify dependencies (including ranges)
x <- str_replace(data$formula, "^\\+", "")
y <- str_extract_all(x, "\\$?[:upper:]{1,2}\\$?[:digit:]+", simplify=TRUE) # matrix (no ranges)
y <- str_extract_all(x, "\\$?[:upper:]{1,2}\\$?[:digit:]+(:\\$?[:upper:]{1,2}\\$?[:digit:]+)?", simplify=TRUE) # matrix (with ranges)
y[is.na(y)] <- ""
data$dep <- do.call(paste, as_tibble(y))
