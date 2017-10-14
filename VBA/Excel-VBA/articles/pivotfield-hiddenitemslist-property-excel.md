---
title: PivotField.HiddenItemsList Property (Excel)
keywords: vbaxl10.chm240129
f1_keywords:
- vbaxl10.chm240129
ms.prod: excel
api_name:
- Excel.PivotField.HiddenItemsList
ms.assetid: 279eeb80-75cd-c758-98b5-668754417482
ms.date: 06/08/2017
---


# PivotField.HiddenItemsList Property (Excel)

Returns or sets a  **Variant** specifying an array of strings that are hidden items for a PivotTable field. Read/write.


## Syntax

 _expression_ . **HiddenItemsList**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The  **HiddenItemsList** property is only valid for Online Analytical Processing (OLAP) data sources; using this property on non-OLAP data sources will return a run-time error.


## Example

The example sets the item list so that only certain items are displayed. It assumes an OLAP PivotTable exists on the active worksheet.


```vb
Sub UseHiddenItemsList() 
 
 ActiveSheet.PivotTables(1).PivotFields(1).HiddenItemsList = _ 
 Array("[Product].[All Products].[Food]", _ 
 "[Product].[All Products].[Drink]") 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

