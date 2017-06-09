---
title: Axis.TickMarkSpacing Property (Excel)
keywords: vbaxl10.chm561102
f1_keywords:
- vbaxl10.chm561102
ms.prod: excel
api_name:
- Excel.Axis.TickMarkSpacing
ms.assetid: 18a23a13-d610-3380-a387-e8f49132dad0
ms.date: 06/08/2017
---


# Axis.TickMarkSpacing Property (Excel)

Returns or sets the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999. Read/write  **Long** .


## Syntax

 _expression_ . **TickMarkSpacing**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Use the  **[MajorUnit](axis-majorunit-property-excel.md)** and **[MinorUnit](axis-minorunit-property-excel.md)** properties to set tick-mark spacing on the value axis.


## Example

This example sets the number of categories between tick marks on the category axis in Chart1.


```vb
Charts("Chart1").Axes(xlCategory).TickMarkSpacing = 10
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

