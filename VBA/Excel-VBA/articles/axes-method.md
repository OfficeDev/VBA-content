---
title: Axes Method
keywords: vbagr10.chm3077608
f1_keywords:
- vbagr10.chm3077608
ms.prod: excel
api_name:
- Excel.Axes
ms.assetid: 040bf3e2-f60f-935b-9803-6f9bf146bee7
ms.date: 06/08/2017
---


# Axes Method

Returns an object that represents either a single axis or a collection of the axes on the chart.

 _expression_. **Axes**( **_Type_**,  **_AxisGroup_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Type** Optional
 **XlAxisType**
. Specifies the axis to return. The reference style of the formula.


|XlAxisType can be one of these XlAxisType constants.|
| **xlValue**|
| **xlCategory** **xlSeriesAxis** (valid only for 3-D charts)|
 **AxisGroup** Optional
 **XlAxisGroup**
. The reference style of the formula.


|XlAxisGroup can be one of these XlAxisGroup constants.|
| **xlPrimary**|
| **xlSecondary**If this argument is omitted, the primary group is used. 3-D charts have only one axis group.|

## Example

This example adds an axis label to the category axis.


```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```

This example turns off major gridlines for the category axis.




```vb
myChart.Axes(xlCategory).HasMajorGridlines = False
```

This example turns off all gridlines for all axes.




```vb
For Each a In myChart.Axes 
 a.HasMajorGridlines = False 
 a.HasMinorGridlines = False 
Next a
```


