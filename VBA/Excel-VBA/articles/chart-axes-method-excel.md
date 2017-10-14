---
title: Chart.Axes Method (Excel)
keywords: vbaxl10.chm149081
f1_keywords:
- vbaxl10.chm149081
ms.prod: excel
api_name:
- Excel.Chart.Axes
ms.assetid: d0520f61-9aff-894b-9975-37dcb5b5fe3c
ms.date: 06/08/2017
---


# Chart.Axes Method (Excel)

Returns an object that represents either a single axis or a collection of the axes on the chart.


## Syntax

 _expression_ . **Axes**( **_Type_** , **_AxisGroup_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|Specifies the axis to return. Can be one of the following  **[XlAxisType](xlaxistype-enumeration-excel.md)** constants: **xlValue** , **xlCategory** , or **xlSeriesAxis** ( **xlSeriesAxis** is valid only for 3-D charts).|
| _AxisGroup_|Optional| **[XlAxisGroup](xlaxisgroup-enumeration-excel.md)**|Specifies the axis group. If this argument is omitted, the primary group is used. 3-D charts have only one axis group.|

### Return Value

Object


## Example

This example adds an axis label to the category axis in Chart1.


```vb
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```

This example turns off major gridlines for the category axis in Chart1.




```vb
Charts("Chart1").Axes(xlCategory).HasMajorGridlines = False
```

This example turns off all gridlines for all axes in Chart1.




```vb
For Each a In Charts("Chart1").Axes 
 a.HasMajorGridlines = False 
 a.HasMinorGridlines = False 
Next a
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

