---
title: PivotField.Calculation Property (Excel)
keywords: vbaxl10.chm240074
f1_keywords:
- vbaxl10.chm240074
ms.prod: excel
api_name:
- Excel.PivotField.Calculation
ms.assetid: abdf0109-da46-1cf6-6f09-c4ba7a3baebd
ms.date: 06/08/2017
---


# PivotField.Calculation Property (Excel)

Returns or sets a  **[XlPivotFieldCalculation](xlpivotfieldcalculation-enumeration-excel.md)** value that represents the type of calculation performed by the specified field. This property is valid only for data fields.


## Syntax

 _expression_ . **Calculation**

 _expression_ A variable that represents a **PivotField** object.


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

