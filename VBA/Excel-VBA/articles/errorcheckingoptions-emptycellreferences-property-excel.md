---
title: ErrorCheckingOptions.EmptyCellReferences Property (Excel)
keywords: vbaxl10.chm698081
f1_keywords:
- vbaxl10.chm698081
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.EmptyCellReferences
ms.assetid: 3d9dd729-8483-aa8e-2d60-312bf3b3e08c
ms.date: 06/08/2017
---


# ErrorCheckingOptions.EmptyCellReferences Property (Excel)

When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells containing formulas that refer to empty cells. **False** disables empty cell reference checking. Read/write **Boolean** .


## Syntax

 _expression_ . **EmptyCellReferences**

 _expression_ A variable that represents an **ErrorCheckingOptions** object.


## Example

In the following example, the  **AutoCorrect Options** button appears for cell A1 which contains a formula that references empty cells.


```vb
Sub CheckEmptyCells() 
 
 Application.ErrorCheckingOptions.EmptyCellReferences = True 
 Range("A1").Formula = "=A2+A3" 
 
End Sub
```


## See also


#### Concepts


[ErrorCheckingOptions Object](errorcheckingoptions-object-excel.md)

