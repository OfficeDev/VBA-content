---
title: PivotTable.DisplayImmediateItems Property (Excel)
keywords: vbaxl10.chm235146
f1_keywords:
- vbaxl10.chm235146
ms.prod: excel
api_name:
- Excel.PivotTable.DisplayImmediateItems
ms.assetid: 796529b1-1f19-4e86-b172-1b2e4173b045
ms.date: 06/08/2017
---


# PivotTable.DisplayImmediateItems Property (Excel)

Returns or sets a  **Boolean** that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty. Set this property to **False** to hide the items in the row and column areas when the data area of the PivotTable is empty. The default value is **True** .


## Syntax

 _expression_ . **DisplayImmediateItems**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example determines how the PivotTable was created and notifies the user. It assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckItemsDisplayed() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine how the PivotTable was created. 
 If pvtTable.DisplayImmediateItems = True Then 
 MsgBox "Fields have been added to the row or column areas for the PivotTable report." 
 Else 
 MsgBox "The PivotTable was created by using object-model calls." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

