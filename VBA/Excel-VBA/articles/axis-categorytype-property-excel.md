---
title: Axis.CategoryType Property (Excel)
keywords: vbaxl10.chm561108
f1_keywords:
- vbaxl10.chm561108
ms.prod: excel
api_name:
- Excel.Axis.CategoryType
ms.assetid: d1e614bb-f560-c65b-7e95-07a997e04861
ms.date: 06/08/2017
---


# Axis.CategoryType Property (Excel)

Returns or sets the category axis type. Read/write  **[XlCategoryType](xlcategorytype-enumeration-excel.md)** .


## Syntax

 _expression_ . **CategoryType**

 _expression_ A variable that represents an **Axis** object.


## Remarks

You cannot set this property for a value axis.


## Example

This example sets the category axis in embedded chart one on worksheet one to use a time scale, with months as the base unit.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

