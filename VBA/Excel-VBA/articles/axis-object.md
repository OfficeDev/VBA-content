---
title: Axis Object
keywords: vbagr10.chm5207088
f1_keywords:
- vbagr10.chm5207088
ms.prod: excel
api_name:
- Excel.Axis
ms.assetid: 708d79de-edcc-ac18-58ec-b9921be9b37e
ms.date: 06/08/2017
---


# Axis Object

Represents a single axis in a chart. The  **Axis** object is a member of the **[Axes](axes-collection.md)** collection.


## Using the Axis Object

Use  **Axes**( _type_,  _group_), where  _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object. _Type_ can be one of the following **XlAxisType** constants: **xlCategory**,  **xlSeries**, or  **xlValue**.  _Group_ can be either of the following **XlAxisGroup** constants: **xlPrimary** or **xlSecondary**. For more information, see the  **[Axes](axes-method.md)** method.

The following example sets the text of the category axis title in the chart.




```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


