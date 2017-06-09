---
title: Name.RefersTo Property (Excel)
keywords: vbaxl10.chm490080
f1_keywords:
- vbaxl10.chm490080
ms.prod: excel
api_name:
- Excel.Name.RefersTo
ms.assetid: 8093e14c-0461-5e49-ef71-16c683044a63
ms.date: 06/08/2017
---


# Name.RefersTo Property (Excel)

Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign. Read/write  **String** .


## Syntax

 _expression_ . **RefersTo**

 _expression_ A variable that represents a **Name** object.


## Example

This example creates a list of all the names in the active workbook, and it shows their formulas in A1-style notation in the language of the macro. The list appears on a new worksheet created by the example.


```vb
Set newSheet = Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.Name 
 newSheet.Cells(i, 2).Value = "'" &; nm.RefersTo 
 i = i + 1 
Next 
newSheet.Columns("A:B").AutoFit
```


## See also


#### Concepts


[Name Object](name-object-excel.md)

