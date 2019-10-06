# excel2r

### Analyse and convert Excel models using R

The objective of this project is to convert large, complex Excel models into R, using R. The approach is to:

1. Read the Excel values and formulae (using the tidyxl package).
2a. Analyse the Excel sheets for consistency (quite useful in itself).
2b. Analyse the formulae to identify scalar, vector and matrix regions (these will become R variables).
3. Translate the Excel formulae to R syntax.
4. Sort the formula into execution order.
5. Write the formula as an R script.

Documentation files:

* project_guide.txt (tba)
* outstanding_issues.txt (tba)
