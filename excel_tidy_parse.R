# parse workbook data
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)
library(tidyxl)

data <- data %>%
	mutate(
		print=case_when(
			is.na(formula) & !is.na(character) ~ character,
			is.na(formula) & !is.na(date) ~ as.character(date),
			is.na(formula) & !is.na(numeric) ~ as.character(numeric),
			is.na(formula) & !is.na(logical) ~ as.character(logical),
			TRUE ~ str_c(address, "=", formula2))
		)
