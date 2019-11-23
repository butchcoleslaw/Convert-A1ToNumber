# Convert-A1ToNumber
 Convert Excel A1 column to column number

This function is used to get a column number when given a column letter.
For example, if you are working with other Excel-related modules that require a column number,
but yet it is easier to know the column letter instead of counting columns, pass to this function the column letter
and this function will return the column number.
Column "A" = 1
Column "B" = 2 ...
Column "N" = 14.
With modern Excel (.xlsx formats), the highest column is "XFD". That will return column number 16384.
NOTE: This function does not limit beyond column "XFD". Actually, you could pass to it column "XFE",
and the value of 16385 would be returned. This function is not limited to the maximum Excel column
(this is by design).  However, there is a warning just in case the limit is exceeded.
