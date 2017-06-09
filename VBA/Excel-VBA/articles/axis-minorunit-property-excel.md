---
title: Axis.MinorUnit Property (Excel)
keywords: vbaxl10.chm561094
f1_keywords:
- vbaxl10.chm561094
ms.prod: excel
api_name:
- Excel.Axis.MinorUnit
ms.assetid: 64cd6523-19c3-7ebc-9b6b-db02667db4d2
ms.date: 06/08/2017
---


# Axis.MinorUnit Property (Excel)

Returns or sets the minor units on the value axis. Read/write  **Double** .


## Syntax

 _expression_ . **MinorUnit**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Setting this property sets the  **[MinorUnitIsAuto](axis-minorunitisauto-property-excel.md)** property to **False** .

Use the  **[TickMarkSpacing](axis-tickmarkspacing-property-excel.md)** property to set tick mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

