# excel crawler functions

# make column codes vector
col <- 1:16384
codes <- if_else(col<=26L, LETTERS[col], paste0(LETTERS[pmax(1, (col-1L) %/% 26L)], LETTERS[(col-1L) %% 26L + 1L]))

# nmake excel address e.g. A3
make_address <- function(row, col){
	paste0(ifelse(col<=26, LETTERS[col],paste0(LETTERS[col %/% 26],LETTERS[col %% 26])), row)
}

# ensnakify
ensnakeify <- function(x) {
	x %>%
		iconv(to="ASCII//TRANSLIT") %>% # remove accents
		str_replace_na() %>% # convert NA to string
		str_to_lower() %>% # convert to lower case
		str_replace_all(pattern="[^[:alnum:]]", replacement=" ") %>% # convert non-alphanumeric to space
		str_trim() %>% # trim leading and trailing spaces
		str_replace_all(pattern="\\s+", replacement="_") # convert remaining spaces to underscore
}

# type of cell
what_type <- function(s){
	n <- as.numeric(str_replace_all(s, ",", ""))
	case_when(
		is.na(s) ~ NA_character_,
		s %in% c("TRUE", "FALSE") ~ "logical",
		is.na(n) ~ "string",
		TRUE ~ "numeric")
}

# find inputs
find_inputs <- function(type, form){
	if (ncol(type)>0){
		inputlist <- vector("list", ncol(type))
		for (col in seq_along(type)){
			row <- which(is.na(form[,col]) & (type[,col] %in% c("numeric", "percentage", "logical")))
			if (length(row)>0){
				inputlist[[col]] <- as.list(make_address(row,col))
			}
		}
		inputlist
	} else {
		list()
	}
}

# avoid as.numeric coercion warnings
as_numeric <- function(x, default=NA_real_){
	suppressWarnings(if_else(is.na(as.numeric(x)), default, as.numeric(x)))
}

# left label

# top label
