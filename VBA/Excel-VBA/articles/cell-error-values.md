---
title: Cell Error Values
keywords: vbaxl10.chm5199688
f1_keywords:
- vbaxl10.chm5199688
ms.prod: excel
ms.assetid: cc4ccabf-37f0-b33d-c03f-13763b85e440
ms.date: 06/08/2017
---


# Cell Error Values

You can insert a cell error value into a cell or test the value of a cell for an error value by using the  **CVErr** function. The cell error values can be one of the following **XlCVError** constants.



|**Constant**|**Error number**|**Cell error value**|
|:-----|:-----|:-----|
| **xlErrDiv0**|2007|#DIV/0!|
| **xlErrNA**|2042|#N/A|
| **xlErrName**|2029|#NAME?|
| **xlErrNull**|2000|#NULL!|
| **xlErrNum**|2036|#NUM!|
| **xlErrRef**|2023|#REF!|
| **xlErrValue**|2015|#VALUE!|

## Example

This example inserts the seven cell error values into cells A1:A7 on Sheet1.


```vb
myArray = Array(xlErrDiv0, xlErrNA, xlErrName, xlErrNull, _ 
 xlErrNum, xlErrRef, xlErrValue) 
For i = 1 To 7 
 Worksheets("Sheet1").Cells(i, 1).Value = CVErr(myArray(i - 1)) 
Next i
```

This example displays a message if the active cell on Sheet1 contains a cell error value. You can use this example as a framework for a cell-error-value error handler.




```vb
Worksheets("Sheet1").Activate 
If IsError(ActiveCell.Value) Then 
 errval = ActiveCell.Value 
 Select Case errval 
 Case CVErr(xlErrDiv0) 
 MsgBox "#DIV/0! error" 
 Case CVErr(xlErrNA) 
 MsgBox "#N/A error" 
 Case CVErr(xlErrName) 
 MsgBox "#NAME? error" 
 Case CVErr(xlErrNull) 
 MsgBox "#NULL! error" 
 Case CVErr(xlErrNum) 
 MsgBox "#NUM! error" 
 Case CVErr(xlErrRef) 
 MsgBox "#REF! error" 
 Case CVErr(xlErrValue) 
 MsgBox "#VALUE! error" 
 Case Else 
 MsgBox "This should never happen!!" 
 End Select 
End If
```


