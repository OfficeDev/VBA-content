---
title: ErrorCheckingOptions.InconsistentFormula Property (Excel)
keywords: vbaxl10.chm698078
f1_keywords:
- vbaxl10.chm698078
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.InconsistentFormula
ms.assetid: 84e482f8-9995-eb26-c4c2-8b258ac1ef9c
ms.date: 06/08/2017
---


# ErrorCheckingOptions.InconsistentFormula Property (Excel)

When set to  **True** (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. **False** disables the inconsistent formula check. Read/write **Boolean** .


## Syntax

 _expression_ . **InconsistentFormula**

 _expression_ A variable that represents an **ErrorCheckingOptions** object.


## Remarks

Consistent formulas in the region must reside to the left and right or above and below the cell containing the inconsistent formula for the  **InconsistentFormula** property to work properly.


## Example

In the following example, when the user selects cell B4 (which contains an inconsistent formula), the  **AutoCorrect Options** button appears.


```vb
Sub CheckFormula() 
 
 Application.ErrorCheckingOptions.InconsistentFormula = True 
 
 Range("A1:A3").Value = 1 
 Range("B1:B3").Value = 2 
 Range("C1:C3").Value = 3 
 
 Range("A4").Formula = "=SUM(A1:A3)" ' Consistent formula. 
 Range("B4").Formula = "=SUM(B1:B2)" ' Inconsistent formula. 
 Range("C4").Formula = "=SUM(C1:C3)" ' Consistent formula. 
 
End Sub
```


## See also


#### Concepts


[ErrorCheckingOptions Object](errorcheckingoptions-object-excel.md)

