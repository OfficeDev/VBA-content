---
title: Axis.CategoryNames Property (Excel)
keywords: vbaxl10.chm561077
f1_keywords:
- vbaxl10.chm561077
ms.prod: excel
api_name:
- Excel.Axis.CategoryNames
ms.assetid: bc565687-ec07-8b60-0bac-a3e13456fefe
ms.date: 06/08/2017
---


# Axis.CategoryNames Property (Excel)

Returns or sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a  **[Range](range-object-excel.md)** object that contains the category names. Read/write **Variant** .


## Syntax

 _expression_ . **CategoryNames**

 _expression_ A variable that represents an **Axis** object.


## Remarks

This property applies only to category axes.


## Example

This example sets the category names for Chart1 to the values in cells B1:B5 on Sheet1.


```vb
Set Charts("Chart1").Axes(xlCategory).CategoryNames = _ 
 Worksheets("Sheet1").Range("B1:B5")
```

This example uses an array to set individual category names for Chart1.




```vb
Charts("Chart1").Axes(xlCategory).CategoryNames = _ 
 Array ("1985", "1986", "1987", "1988", "1989")
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

