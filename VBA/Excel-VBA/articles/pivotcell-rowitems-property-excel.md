---
title: PivotCell.RowItems Property (Excel)
keywords: vbaxl10.chm692078
f1_keywords:
- vbaxl10.chm692078
ms.prod: excel
api_name:
- Excel.PivotCell.RowItems
ms.assetid: 4833f772-9abd-a2fa-e3f0-e86f54caf05e
ms.date: 06/08/2017
---


# PivotCell.RowItems Property (Excel)

Returns a  **[PivotItemList](pivotitemlist-object-excel.md)** collection that corresponds to the items on the category axis that represent the selected cell.


## Syntax

 _expression_ . **RowItems**

 _expression_ A variable that represents a **PivotCell** object.


## Example

This example determines if the data item in cell B5 is under the Inventory item in the first row field and notifies the user. The example assumes a PivotTable exists on the active worksheet and that column B of the worksheet contains a row item of the PivotTable.


```vb
Sub CheckRowItems() 
 
 ' Determine if there is a match between the item and row field. 
 If Application.Range("B5").PivotCell.RowItems.Item(1) = "Inventory" Then 
 MsgBox "Cell B5 is a member of the 'Inventory' row field. 
 Else 
 MsgBox "Cell B5 is not a member of the 'Inventory' row field. 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

