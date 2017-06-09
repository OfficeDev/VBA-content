---
title: Name.RefersToR1C1 Property (Excel)
keywords: vbaxl10.chm490086
f1_keywords:
- vbaxl10.chm490086
ms.prod: excel
api_name:
- Excel.Name.RefersToR1C1
ms.assetid: 6661dc25-44cd-ac43-9347-93ed7583c9b1
ms.date: 06/08/2017
---


# Name.RefersToR1C1 Property (Excel)

Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign. Read/write  **String** .


## Syntax

 _expression_ . **RefersToR1C1**

 _expression_ A variable that represents a **Name** object.


## Example

This example creates a new worksheet and then inserts a list of all the names in the active workbook, including their formulas (in R1C1-style notation and in the language of the macro).


```vb
Set newSheet = ActiveWorkbook.Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.Name 
 newSheet.Cells(i, 2).Value = "'" &; nm.RefersToR1C1 
 i = i + 1 
Next
```


## See also


#### Concepts


[Name Object](name-object-excel.md)

