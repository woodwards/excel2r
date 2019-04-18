# test tidyxl

library(tidyxl)

wbnames <- c("Simplified forecaster.xlsx")
wbnames <- c("Copy of NBO 2019 _Final_withGHG.xlsx")

value <- xlsx_cells(wbnames[1])

value$character_formatted[1]

form <- value$formula[444]
parsed <- xlex(form)
form <- value$formula[1449]
parsed <- xlex(form)
setNames(parsed$token, parsed$type)
