---
title: CellFormat.Interior Property (Excel)
keywords: vbaxl10.chm676075
f1_keywords:
- vbaxl10.chm676075
ms.prod: excel
api_name:
- Excel.CellFormat.Interior
ms.assetid: aa11d693-0713-1f0c-0ef0-87bb81f705bd
ms.date: 06/08/2017
---


# CellFormat.Interior Property (Excel)

Returns an  **[Interior](interior-object-excel.md)** object allowing the user to set or return the search criteria based on the cell's interior format.


## Syntax

 _expression_ . **Interior**

 _expression_ A variable that represents a **CellFormat** object.


## Example

This example sets the search criteria to identify cells that contain a solid yellow interior, creates a cell with this condition, finds this cell, and notifies the user.


```vb
Sub SearchCellFormat() 
 
 ' Set the search criteria for the interior of the cell format. 
 With Application.FindFormat.Interior 
 .ColorIndex = 6 
 .Pattern = xlSolid 
 .PatternColorIndex = xlAutomatic 
 End With 
 
 ' Create a yellow interior for cell A5. 
 Range("A5").Select 
 With Selection.Interior 
 .ColorIndex = 6 
 .Pattern = xlSolid 
 .PatternColorIndex = xlAutomatic 
 End With 
 Range("A1").Select 
 MsgBox "Cell A5 has a yellow interior." 
 
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

