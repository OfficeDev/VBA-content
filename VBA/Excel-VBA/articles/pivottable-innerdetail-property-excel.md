---
title: PivotTable.InnerDetail Property (Excel)
keywords: vbaxl10.chm235084
f1_keywords:
- vbaxl10.chm235084
ms.prod: excel
api_name:
- Excel.PivotTable.InnerDetail
ms.assetid: 385449ab-fbe2-8b69-374e-a5d374a3f76f
ms.date: 06/08/2017
---


# PivotTable.InnerDetail Property (Excel)

Returns or sets the name of the field that will be shown as detail when the  **ShowDetail** property is **True** for the innermost row or column field. Read/write **String** .


## Syntax

 _expression_ . **InnerDetail**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

This property isn't available for OLAP data sources.


## Example

This example displays the name of the field that will be shown as detail when the  **ShowDetail** property is **True** for the innermost row field or column field.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox pvtTable.InnerDetail
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

