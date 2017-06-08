---
title: MinorUnit Property
keywords: vbagr10.chm3077551
f1_keywords:
- vbagr10.chm3077551
ms.prod: excel
api_name:
- Excel.MinorUnit
ms.assetid: 9da86e1c-dfc2-49c8-e6bd-1e5529b2da33
ms.date: 06/08/2017
---


# MinorUnit Property

Returns or sets the minor units on the axis. Read/write Double.

 _expression_. **MinorUnit**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

Setting this property sets the  **[MinorUnitIsAuto](minorunitisauto-property.md)** property to  **False**.

Use the  **[TickMarkSpacing](tickmarkspacing-property.md)** property to set tick-mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```


