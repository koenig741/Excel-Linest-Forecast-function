# Excel-Linest-Forecast-function
If you use linest for linear regression, this function will save time creating the model forecast

Description:
This Excel Visual Basic function is designed for use with regressions using the LINEST function.  The function allows the user to select 
ranges for the regression constants and the model variables (features).  

Usage:
The formula is used "=LmFcast(cell range for constants, cell range for variables)".  In a cell enter "=LmFcast(" then Control A ... this pops up a dialog box to help you select the variables and constants correctly.  Make sure to fix the constant cell range using "$".  The attached worksheet contains the code and sample to demonstrate usage and a math check.  

Function:
The function code use the first element of the constants as the intercept then adds each pair (in opposite directions) until the remainder of the constants and variables are multiplied and added.  

```
Function LmFcast(rConstants As Range, rVariables As Range)
  Dim ArrV() As Variant: ArrV = Application.WorksheetFunction.Transpose(rVariables)
  Dim ArrC() As Variant: ArrC = Application.WorksheetFunction.Transpose(rConstants)
  Dim C As Long
  Dim V As Long: V = 1
  LmFcast = ArrC(UBound(ArrC, 1), 1)
  For C = UBound(ArrC, 1) - 1 To 1 Step -1  'First array dimension is rows.
      LmFcast = LmFcast + ArrC(C, 1) * ArrV(V, 1)
      V = V + 1
  Next C
End Function
```
