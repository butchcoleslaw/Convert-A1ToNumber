Function Convert-A1ToNumber {
<#
.SYNOPSIS
Converts an Excel Column Letter (A1 format) to a Column Number

.DESCRIPTION
This function is used to get a column number when given a column letter.
For example, if you are working with other Excel-related modules that require a column number,
but yet it is easier to know the column letter instead, pass to this function the column letter
and this function will return the column number.
Column "A" = 1
Column "B" = 2 ...
Column "N" = 14.
With modern Excel (.xlsx formats), the highest column is "XFD". That will return column number 16384.
NOTE: This function does not limit beyond column "XFD". Actually, you could pass to it column "XFE",
and the value of 16385 would be returned. This function is not limited to the maximum Excel column
(this is by design).  However, there is a warning just in case the limit is exceeded.

.PARAMETER ColName
Mandatory. Any sequence of letters between "A" and "XFD".

.INPUTS
None - other than parameters above

.OUTPUTS
Integer value that represents the ordinal value of the column letter given.

.NOTES
Version:        1.0
Author:         Ken Friddle  listosystems@gmail.com
Creation Date:  2019/11/19
Purpose/Change: Initial function development
  
.EXAMPLE
$ColumnNumber = Convert-A1ToNumber "N"
$ColumnNumber

14

#>
[CmdletBinding()]

Param([parameter(Mandatory=$true)]
    [string]$ColName)

    [int]$value = 0

    $ColName = $ColName.ToUpper()
    $ColNameArray = $ColName.ToCharArray()
    [array]::Reverse($ColNameArray)
    for($i = 0; $i -le $ColNameArray.Length - 1;$i++) {
        if ($i -gt 0) {
            $value += ([math]::Pow(26,$i)) * ([byte][char]$ColNameArray[$i] - 64)
        } else {
            $value += ([byte][char]$ColNameArray[$i] - 64)
        }
    }
    if ($value -gt 16384) {
        Write-Warning "Input column letters do not exist in Excel."
        Write-Warning "Value returned is not valid for Excel."
    }
    Return $value
} #End function
