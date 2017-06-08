---
title: Axes Collection
keywords: vbagr10.chm131099
f1_keywords:
- vbagr10.chm131099
ms.prod: excel
api_name:
- Excel.Axes
ms.assetid: 89ebeb9d-3c16-0bb0-35a8-9a07483c4eb6
ms.date: 06/08/2017
---


# Axes Collection

A collection of all the  **[Axis](axis-object.md)** objects in the specified chart.


## Using the Axes Collection

Use the  **Axes** method to return the **Axes** collection. The following example displays the number of axes in the chart.


```vb
With myChart 
 MsgBox .Axes.Count 
End With
```

Use  **Axes**( _type_,  _group_), where  _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object.

 _Type_ can be one of the following **XlAxisType**constants.



|XlAxisType can be one of these XlAxisType constants.|
| **xlCategory**|
| **xlSeries** **xlValue**|
 _Group_ can be either of the following **XlAxisGroup** constants: **xlPrimary** or **xlSecondary**. For more information, see the  **[Axes](axes-method.md)** method. 

The following example sets the title text for the category axis.




```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


