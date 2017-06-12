---
title: PivotCell.PivotItem Property (Excel)
keywords: vbaxl10.chm692077
f1_keywords:
- vbaxl10.chm692077
ms.prod: excel
api_name:
- Excel.PivotCell.PivotItem
ms.assetid: 3b131e96-8589-9d72-d4d9-afe2d3d6137c
ms.date: 06/08/2017
---


# PivotCell.PivotItem Property (Excel)

Returns a  **[PivotItem](pivotitem-object-excel.md)** object that represents the PivotTable item containing the upper-left corner of the specified range.


## Syntax

 _expression_ . **PivotItem**

 _expression_ A variable that represents a **PivotCell** object.


## Example

This example displays the name of the PivotTable item that contains the active cell on Sheet1.


```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the item " &; _ 
 ActiveCell.PivotItem.Name
```


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

