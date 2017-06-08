---
title: PivotItem.ParentItem Property (Excel)
keywords: vbaxl10.chm246079
f1_keywords:
- vbaxl10.chm246079
ms.prod: excel
api_name:
- Excel.PivotItem.ParentItem
ms.assetid: 7d0959e5-5abc-c84f-7037-19b761f36294
ms.date: 06/08/2017
---


# PivotItem.ParentItem Property (Excel)

Returns a  **PivotItem** object that represents the parent PivotTable item in the parent **[PivotField](pivotfield-object-excel.md)** object (the field must be grouped so that it has a parent). Read-only.


## Syntax

 _expression_ . **ParentItem**

 _expression_ A variable that represents a **PivotItem** object.


## Remarks

This property isn't available for OLAP data sources.


## Example

This example displays the name of the parent item for the item that contains the active cell.


```vb
Worksheets("Sheet1").Activate 
MsgBox "This item is a subitem of " &; _ 
 ActiveCell.PivotItem.ParentItem.Name
```


## See also


#### Concepts


[PivotItem Object](pivotitem-object-excel.md)

