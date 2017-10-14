---
title: CellFormat.Font Property (Excel)
keywords: vbaxl10.chm676074
f1_keywords:
- vbaxl10.chm676074
ms.prod: excel
api_name:
- Excel.CellFormat.Font
ms.assetid: 2a0ee538-e7fa-581f-4c8b-b48e61b46f8a
ms.date: 06/08/2017
---


# CellFormat.Font Property (Excel)

Returns a  **[Font](font-object-excel.md)** object, allowing the user to set or return the search criteria based on the cell's font format.


## Syntax

 _expression_ . **Font**

 _expression_ A variable that represents a **CellFormat** object.


## Example

This example sets the search criteria to identify cells that contain red font, creates a cell with this condition, finds this cell, and notifies the user.


```vb
Sub SearchCellFormat() 
 
 ' Set the search criteria for the font of the cell format. 
 Application.FindFormat.Font.ColorIndex = 3 
 
 ' Set the color index of the font for cell A5 to red. 
 Range("A5").Font.ColorIndex = 3 
 Range("A5").Formula = "Red font" 
 Range("A1").Select 
 MsgBox "Cell A5 has red font" 
 
 ' Find the cells based on the search criteria. 
 Cells.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _ 
 xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _ 
 , SearchFormat:=True).Activate 
 
 MsgBox "Microsoft Excel has found this cell matching the search criteria." 
 
End Sub
```


## See also


#### Concepts


[CellFormat Object](cellformat-object-excel.md)

