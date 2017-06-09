---
title: PivotItem.Position Property (Excel)
keywords: vbaxl10.chm246081
f1_keywords:
- vbaxl10.chm246081
ms.prod: excel
api_name:
- Excel.PivotItem.Position
ms.assetid: 07e78622-f869-40d0-276a-b015ebe7a90f
ms.date: 06/08/2017
---


# PivotItem.Position Property (Excel)

Returns or sets a  **Long** value that represents the position of the item in its field, if the item is currently showing.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **PivotItem** object.


## Example

This example displays the position number of the PivotTable item that contains the active cell.


```vb
Worksheets("Sheet1").Activate 
MsgBox "The active item is in position number " &; _ 
 ActiveCell.PivotItem.Position
```


## See also


#### Concepts


[PivotItem Object](pivotitem-object-excel.md)

