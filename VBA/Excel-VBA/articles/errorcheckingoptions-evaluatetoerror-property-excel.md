---
title: ErrorCheckingOptions.EvaluateToError Property (Excel)
keywords: vbaxl10.chm698075
f1_keywords:
- vbaxl10.chm698075
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.EvaluateToError
ms.assetid: f6a7c606-6da6-defd-9ca5-9ce46805e2d7
ms.date: 06/08/2017
---


# ErrorCheckingOptions.EvaluateToError Property (Excel)

When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells that contain formulas evaluating to an error. **False** disables error checking for cells that evaluate to an error value. Read/write **Boolean** .


## Syntax

 _expression_ . **EvaluateToError**

 _expression_ A variable that represents an **ErrorCheckingOptions** object.


## Example

In the following example, the AutoCorrect Options button appears for cell A3, which contains a divide-by-zero error.


```vb
Sub CheckEvaluationError() 
 
 ' Simulate a divide-by-zero error. 
 Application.ErrorCheckingOptions.EvaluateToError = True 
 Range("A1").Value = 1 
 Range("A2").Value = 0 
 Range("A3").Formula = "=A1/A2" 
 
End Sub
```


## See also


#### Concepts


[ErrorCheckingOptions Object](errorcheckingoptions-object-excel.md)

