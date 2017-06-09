---
title: PivotField.BaseItem Property (Excel)
keywords: vbaxl10.chm240096
f1_keywords:
- vbaxl10.chm240096
ms.prod: excel
api_name:
- Excel.PivotField.BaseItem
ms.assetid: 11561507-043a-2b64-1b60-3cdbd93a656c
ms.date: 06/08/2017
---


# PivotField.BaseItem Property (Excel)

Returns or sets the item in the base field for a custom calculation. Valid only for data fields. Read/write  **Variant** .


## Syntax

 _expression_ . **BaseItem**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

This property is not available for OLAP data sources.


## Example

This example sets the data field in the PivotTable report on Sheet1 to calculate the difference from the base field, sets the base field to the field named "ORDER_DATE," and then sets the base item to the item named "5/16/89."


```vb
With Worksheets("Sheet1").Range("A3").PivotField 
 .Calculation = xlDifferenceFrom 
 .BaseField = "ORDER_DATE" 
 .BaseItem = "5/16/89" 
End With
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

