---
title: MajorUnit Property
keywords: vbagr10.chm5207641
f1_keywords:
- vbagr10.chm5207641
ms.prod: excel
api_name:
- Excel.MajorUnit
ms.assetid: 46d4d4e0-f285-2800-f539-72e7acb98948
ms.date: 06/08/2017
---


# MajorUnit Property

Returns or sets the major units for the axis. Read/write  **Double**.


## Remarks

Setting this property sets the  **[MajorUnitIsAuto](majorunitisauto-property.md)** property to  **False**.

Use the  **[TickMarkSpacing](tickmarkspacing-property.md)** property to set tick-mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```


