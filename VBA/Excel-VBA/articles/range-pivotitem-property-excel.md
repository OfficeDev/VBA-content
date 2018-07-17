---
title: Range.PivotItem Property (Excel)
keywords: vbaxl10.chm144176
f1_keywords:
- vbaxl10.chm144176
ms.prod: excel
api_name:
- Excel.Range.PivotItem
ms.assetid: 02a41786-074b-ae34-5d2c-407006fe526d
ms.date: 06/08/2017
---


# Range.PivotItem Property (Excel)

Returns a  **[PivotItem](pivotitem-object-excel.md)** object that represents the PivotTable item containing the upper-left corner of the specified range.


## Syntax

 _expression_ . **PivotItem**

 _expression_ A variable that represents a **Range** object.


## Example

This example displays the name of the PivotTable item that contains the active cell on Sheet1.


```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the item " &; _ 
 ActiveCell.PivotItem.Name
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

