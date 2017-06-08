---
title: Axes Object (Excel)
keywords: vbaxl10.chm571072
f1_keywords:
- vbaxl10.chm571072
ms.prod: excel
api_name:
- Excel.Axes
ms.assetid: 581e51e5-3dbb-7f0c-a87d-2d44f67dad0b
ms.date: 06/08/2017
---


# Axes Object (Excel)

A collection of all the  **[Axis](axis-object-excel.md)** objects in the specified chart.


## Remarks

Use the  **Axes** method to return the **Axes** collection.

Use  **Axes** ( _type_, _group_ ), where _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](xlaxistype-enumeration-excel.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _Group_ can be one of the following **[XlAxisGroup](xlaxisgroup-enumeration-excel.md)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](chart-axes-method-excel.md)** method.


## Example

The following example displays the number of axes on embedded chart one on worksheet one.


```
With Worksheets(1).ChartObjects(1).Chart 
 MsgBox.Axes.Count 
End With
```

The following example sets the category axis title text on the chart sheet named "Chart1."




```
With Charts("chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


## Methods



|**Name**|
|:-----|
|[Item](axes-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](axes-application-property-excel.md)|
|[Count](axes-count-property-excel.md)|
|[Creator](axes-creator-property-excel.md)|
|[Parent](axes-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
