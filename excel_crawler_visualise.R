# visualise workbook data
# Simon Woodward, DairyNZ 2019

# load libraries
library(tidyverse)
library(digest)

# load functions
source("excel_crawler_functions.R")

# make column codes vector
xlcol <- 1:10000 # or as wide as necessary, max is 16384
xlcode <- if_else(xlcol<=26L, LETTERS[xlcol], paste0(LETTERS[pmax(1, (xlcol-1L) %/% 26L)], LETTERS[(xlcol-1L) %% 26L + 1L]))
names(xlcol) <- xlcode

#### normalise numeric formats ####
print(paste("Visualising workbook", wbname))
print("Normalising numeric values")
x <- str_replace_all(data$value, "[,\\s]", "") # strip commas and spaces
data$numeric <- as_numeric(x)
data$percent <- as_numeric(str_replace(str_match(x, "^-?[:digit:]+(\\.[:digit:]+)?%$")[,1], "%", ""))/100
data$currency <- as_numeric(str_replace_all(str_match(x, "^-?\\$[:digit:]+(\\.[:digit:]+)?$")[,1], "\\$", ""))
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
# note these regex are intentionally simplified but will handle most sensible cases
range <- "\\$?[A-Z]+\\$?[0-9]+(:\\$?[A-Z]+\\$?[0-9]+)?"
sheet <- "[a-zA-Z][a-zA-Z0-9\\s\\+\\-\\&\\_\\(\\)]*" # add any needed sheetname characters
book <- "\\[[a-zA-Z0-9][a-zA-Z0-9\\s\\+\\-\\&\\_\\.\\:\\\\]*\\]" # add any needed sheetname/bookname characters
pattern <- str_c("(('?((", book, ")?", sheet, ")'?!)?", range, ")(?!\\()") # following ( indicates function name
# find all references
print("Identifying cell references")
data <- data %>%
	select(-matches("formula2"), # remove old result if it exists
		   -starts_with("V", ignore.case=FALSE)) %>% # remove any old V columns
	mutate(i=1:n()) %>% # add row numbers
	bind_cols(as_tibble( # new columns are named V1, V2 ...
		str_extract_all(.$formula, pattern, simplify=TRUE) # returns a matrix
						))
# gather references for cleaning
x <- data %>%
	select(i, wbnum, wsname, col, row, formula, starts_with("V", ignore.case=FALSE)) %>%
	gather(pos, ref0, starts_with("V", ignore.case=FALSE)) %>%
	filter(ref0>"") %>%
	arrange(i) %>%
	mutate(ref=str_extract(ref0, range))
# parse. this is slow. about 8 minutes!!! and hard to debug
# x <- x %>%
# 	mutate(
# 		# calculate relative coordinates (can be partial relatively)
# 		col1=xlcol[str_match(ref, "^[A-Z]+")] - col,
# 		row1=as.integer(str_match(ref, "(?<=^\\$?[A-Z]{1,3})[0-9]+")) - row,
# 		col2=xlcol[str_match(ref, "(?<=:)[A-Z]+")] - col,
# 		row2=as.integer(str_match(ref, "(?<=:\\$?[A-Z]{1,3})[0-9]+")) - row,
# 		# construct new range (put column first then row)
# 		ref1a=str_c(
# 			if_else(!is.na(col1), str_c("C[", col1, "]"),
# 					str_extract(ref, "^\\$[A-Z]+")),
# 			if_else(!is.na(row1), str_c("R[", row1, "]"),
# 					str_extract(ref, "(?<=^\\$?[A-Z]{1,3})\\$[0-9]+"))),
# 		ref1b=str_c(
# 			if_else(!is.na(col2), str_c(":C[", col2, "]"),
# 					str_extract(ref, ":\\$[A-Z]+")),
# 			if_else(!is.na(row2), str_c("R[", row2, "]"),
# 					str_extract(ref, "(?<=:\\$?[A-Z]{1,3})\\$[0-9]+"))),
# 		ref1=if_else(is.na(ref1b), ref1a, str_c(ref1a, ref1b)),
# 		# construct cleaned reference
# 		ref2=str_replace(ref0, range, ref1),
# 		ref2=str_replace_all(ref2, "'", ""),
# 		ref2=str_replace_all(ref2, "^[^!]*$", str_c(wsname, "!", ref2)),
# 		ref2=str_replace_all(ref2, "^(?!\\[)", "[0]")
# 	)
# calculate relative coordinates (can be partial relatively)
print("Relativising cell references")
x$col1=str_match(x$ref, "^[A-Z]+")[,1]
i <- !is.na(x$col1)
x$col1[i]=xlcol[x$col1[i]] - x$col[i]
x$row1=str_match(x$ref, "(?<=^\\$?[A-Z]{1,3})[0-9]+")[,1]
i <- !is.na(x$row1)
x$row1[i]=as.integer(x$row1[i]) - x$row[i]
x$col2=str_match(x$ref, "(?<=:)[A-Z]+")[,1]
i <- !is.na(x$col2)
x$col2[i]=xlcol[x$col2[i]] - x$col[i]
x$row2=str_match(x$ref, "(?<=:\\$?[A-Z]{1,3})[0-9]+")[,1]
i <- !is.na(x$row2)
x$row2[i]=as.integer(x$row2[i]) - x$row[i]
# construct new range (put column first then row)
x$ref1a=str_c(
	if_else(!is.na(x$col1), str_c("C[", x$col1, "]"),
			str_extract(x$ref, "^\\$[A-Z]+")),
	if_else(!is.na(x$row1), str_c("R[", x$row1, "]"),
			str_extract(x$ref, "(?<=^\\$?[A-Z]{1,3})\\$[0-9]+")))
x$ref1b=str_c(
	if_else(!is.na(x$col2), str_c(":C[", x$col2, "]"),
			str_extract(x$ref, ":\\$[A-Z]+")),
	if_else(!is.na(x$row2), str_c("R[", x$row2, "]"),
			str_extract(x$ref, "(?<=:\\$?[A-Z]{1,3})\\$[0-9]+")))
x$ref1=if_else(is.na(x$ref1b), x$ref1a, str_c(x$ref1a, x$ref1b))
# construct cleaned reference
x$ref2=str_replace(x$ref0, fixed(x$ref), x$ref1) # very slow if we don't use literal
x$ref2=str_replace_all(x$ref2, "'", "")
i=!str_detect(x$ref2, "!")
x$ref2[i]=str_c(x$wsname[i], "!", x$ref2[i])
i=!str_detect(x$ref2, "^\\[")
x$ref2[i]=str_c("[0]", x$ref2[i])
# insert cleaned references into formula
y <- x %>%
	select(i, formula, pos, ref2) %>%
	spread(pos, ref2) %>%
	mutate(
		formula2=str_replace(formula, "^\\+", ""), # remove leading +
		formula2=str_replace_all(formula2, pattern, mysep) # use a placeholder
		)
ii <- str_extract(names(y), "^V[0-9]+") # get V column names
ii <- ii[!is.na(ii)]
i <- ii[1]
for (i in ii){
	j <- !is.na(y[[i]]) # avoid NA
	y$formula2[j] <- str_replace(y$formula2[j], fixed(mysep), y[[i]][j])
}
# join back into data
print("Rebuilding formulae")
data <- data %>%
	left_join(
		select(y, i, formula2),
		by=c("i")
		) %>%
	arrange(wbnum, wsname, col, row)

# anonymise to identify regions of identical formulae
# https://stackoverflow.com/questions/27442991/how-to-get-a-hash-code-as-integer-in-r
data <- data %>%
	mutate(
		hash=case_when(
			!is.na(formula2) ~ map_chr(formula2, digest, algo="xxhash32"),
			# !is.na(value2) ~ map_chr(formula2, digest, algo="xxhash32"),
			TRUE ~ NA_character_
			),
		hashi=strtoi(str_sub(hash, end=-2), 16) # 8 hex digits can give 32-bit integer overflow
	)

#### plots ####
print("Plotting visualisation")
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
fname <- paste0(str_replace_all(data$wbname[1], " |\\.", "_"), "_formulaplot.png")
png(fname, width=297*1.3, heigh=210, units="mm", res=300)
print(p)
dev.off()

# numerical value plot
p <- ggplot() +
	labs(title=data$wbname[1], fill="Value#") +
	geom_raster(data=data,
				mapping=aes(x=col,
							y=row,
							# colour=atan(value2),
							# size=str_length(formula2),
							# fill=log10(pmax(abs(value2), 1e-3))
							fill=rank(value2, na.last="keep", ties.method="min")
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
fname <- paste0(str_replace_all(data$wbname[1], " |\\.", "_"), "_valueplot.png")
png(fname, width=297*1.3, heigh=210, units="mm", res=300)
print(p)
dev.off()

