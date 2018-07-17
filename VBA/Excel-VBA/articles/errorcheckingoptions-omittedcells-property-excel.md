---
title: ErrorCheckingOptions.OmittedCells Property (Excel)
keywords: vbaxl10.chm698079
f1_keywords:
- vbaxl10.chm698079
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.OmittedCells
ms.assetid: a337da5d-4f02-d24c-c59a-288b4a9c9117
ms.date: 06/08/2017
---


# ErrorCheckingOptions.OmittedCells Property (Excel)

When set to  **True** (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. **False** disables error checking for omitted cells. Read/write **Boolean** .


## Syntax

 _expression_ . **OmittedCells**

 _expression_ A variable that represents an **ErrorCheckingOptions** object.


## Example

In the following example, the  **AutoCorrect Options** button appears for cell A4, which contains a formula.


```vb
Sub CheckOmittedCells() 
 
 Application.ErrorCheckingOptions.OmittedCells = True 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Value = 3 
 Range("A4").Formula = "=Sum(A1:A2)" 
 
End Sub
```


## See also


#### Concepts


[ErrorCheckingOptions Object](errorcheckingoptions-object-excel.md)

