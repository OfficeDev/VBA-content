---
title: PivotTable.ShowPageMultipleItemLabel Property (Excel)
keywords: vbaxl10.chm235150
f1_keywords:
- vbaxl10.chm235150
ms.prod: excel
api_name:
- Excel.PivotTable.ShowPageMultipleItemLabel
ms.assetid: 2f816331-4017-a208-d1b2-fea219d2ca71
ms.date: 06/08/2017
---


# PivotTable.ShowPageMultipleItemLabel Property (Excel)

When set to  **True** (default), "(Multiple Items)" will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowPageMultipleItemLabel**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example determines if "(Multiple Items)" will be displayed in the PivotTable cell and notifies the user. The example assumes that a PivotTable exists on the active worksheet.


```vb
Sub UseShowPageMultipleItemLabel() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if multiple items are allowed. 
 If pvtTable.ShowPageMultipleItemLabel = True Then 
 MsgBox "The words 'Multiple Items' can be displayed." 
 Else 
 MsgBox "The words 'Multiple Items' cannot be displayed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

