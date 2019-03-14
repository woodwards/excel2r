# visualise workbook data
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)
library(digest)

# load functions
source("excel_crawler_functions.R")

#### normalise numeric formats ####
x <- str_replace_all(data$value, "[,\\s]", "") # strip commas and spaces
# numbers -1.0E5, -2, 2,000.000
data$numeric <- as_numeric(x)
# percentages -5.0%, -5%
data$percent <- as_numeric(str_replace(str_match(x, "^-?[:digit:]+(\\.[:digit:]+)?%$")[,1], "%", ""))/100
# currency -$ 1,000.00
data$currency <- as_numeric(str_replace_all(str_match(x, "^-?\\$[:digit:]+(\\.[:digit:]+)?$")[,1], "\\$", ""))
# dates and times
# data$datetime <-
# combine values
data <- data %>%
	mutate(
		value2=case_when(
			!is.na(numeric) ~ numeric,
			!is.na(percent) ~ percent,
			!is.na(currency) ~ currency,
			TRUE ~ NA_real_)
	)

#### normalise formulae ####
# note these regex are intentionally simplified but will handle most cases
range <- "\\$?[A-Z]+\\$?[0-9]+(:\\$?[A-Z]+\\$?[0-9]+)?"
sheet <- "[a-zA-Z][a-zA-Z0-9\\s\\+\\-\\&\\_\\(\\)]*" # add any needed sheetname characters
book <- "\\[[a-zA-Z0-9][a-zA-Z0-9\\s\\+\\-\\&\\_\\.\\:\\\\]*\\]" # add any needed sheetname/bookname characters
pattern <- str_c("(('?((", book, ")?", sheet, ")'?!)?", range, ")(?!\\()") # following ( indicates function name
# find all references
# y <- str_extract_all(data$formula, pattern, simplify=TRUE) # returns a matrix
# y[is.na(y)] <- ""
# data$refs <- do.call(paste, as_tibble(y))
# # normalise
# for (i in 1:nrow(y)){
# 	for (j in 1:ncol(y)){
# 		if (y[i,j]>""){
# 			y[i,j] <- str_replace_all(y[i,j], "'", "")
# 			y[i,j] <- str_replace_all(y[i,j], "^[^!]*$", paste0(data$wsname[i], "!", y[i,j]))
# 			y[i,j] <- str_replace_all(y[i,j], "^(?!\\[)", "[0]")
# 		}
# 	}
# }
# data$refs2 <- do.call(paste, as_tibble(y))
# anonymise to identify regions of identical formulae (very rough)
# https://stackoverflow.com/questions/27442991/how-to-get-a-hash-code-as-integer-in-r
# https://www.jimhester.com/2018/04/12/vectorize/
range <- "\\$?[A-Z]+\\$?[0-9]+(:\\$?[A-Z]+\\$?[0-9]+)?" # have to include $ in case partial row/col lock
pattern <- str_c("(", range, ")(?!\\()") # following ( indicates function name
data <- data %>%
	mutate(
		formula2=str_replace_all(formula, pattern, "REF"),
		hash=case_when(
			!is.na(formula2) ~ map_chr(formula2, digest, algo="xxhash32"),
			!is.na(value2) ~ map_chr(formula2, digest, algo="xxhash32"),
			TRUE ~ NA_character_
			),
		hashi=strtoi(str_sub(hash, end=-2), 16) # 8 hex digits can give 32-bit integer overflow
	)

#### plots ####
library(viridis)
library(ggthemes)
p <- ggplot() +
	labs(title=data$wbname[1], fill="Formula#") +
	geom_raster(data=data,
			   mapping=aes(x=col,
			   			y=row,
						# colour=atan(value2),
						# size=str_length(formula2),
						fill=hashi
						# alpha=log(abs(value2)+1)
						)
			   # ,
			   # size=0.1
			   ) +
	scale_y_reverse() +
	# scale_colour_gradientn(colours=terrain.colors(16)) +
	# theme_few() +
	scale_fill_viridis() +
	facet_wrap( ~ wsname, scales="free")
print(p)
fname <- paste0("facetplot_", str_replace_all(data$wbname[1], " |\\.", "_"), ".png")
png(fname, width=297*1.3, heigh=210, units="mm", res=300)
print(p)
dev.off()


