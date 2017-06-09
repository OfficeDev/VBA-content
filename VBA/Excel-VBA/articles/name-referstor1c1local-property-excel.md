---
title: Name.RefersToR1C1Local Property (Excel)
keywords: vbaxl10.chm490087
f1_keywords:
- vbaxl10.chm490087
ms.prod: excel
api_name:
- Excel.Name.RefersToR1C1Local
ms.assetid: 314b8764-5f5c-9a2f-87a7-54637de59bbd
ms.date: 06/08/2017
---


# Name.RefersToR1C1Local Property (Excel)

Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign. Read/write  **String** .


## Syntax

 _expression_ . **RefersToR1C1Local**

 _expression_ A variable that represents a **Name** object.


## Example

This example creates a new worksheet and then inserts a list of all the names in the active workbook, including their formulas (in R1C1-style notation and in the language of the user).


```vb
Set newSheet = ActiveWorkbook.Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.NameLocal 
 newSheet.Cells(i, 2).Value = "'" &; nm.RefersToR1C1Local 
 i = i + 1 
Next
```


## See also


#### Concepts


[Name Object](name-object-excel.md)

