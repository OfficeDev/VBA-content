---
title: Chart.HasAxis Property (Excel)
keywords: vbaxl10.chm149113
f1_keywords:
- vbaxl10.chm149113
ms.prod: excel
api_name:
- Excel.Chart.HasAxis
ms.assetid: f2df9f16-980d-fd02-3e09-6d6903dbb6c6
ms.date: 06/08/2017
---


# Chart.HasAxis Property (Excel)

Returns or sets which axes exist on the chart. Read/write  **Variant** .


## Syntax

 _expression_ . **HasAxis**( **_Index1_** , **_Index2_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index1_|Required| **Variant**|The axis type. Series axes apply only to 3-D charts. Can be one of the  **[XlAxisType](xlaxistype-enumeration-excel.md)** constants.|
| _Index2_|Optional| **Variant**|The axis group. 3-D charts have only one set of axes. Can be one of the  **[XlAxisGroup](xlaxisgroup-enumeration-excel.md)** constants.|

## Remarks

You must enter a value for at least one of the parameters when setting this property.

Microsoft Excel may create or delete axes if you change the chart type or the  **[Axis.AxisGroup](axis-axisgroup-property-excel.md)** , **[Chart.AxisGroup](chartgroup-axisgroup-property-excel.md)** , or **[Series.AxisGroup](series-axisgroup-property-excel.md)** properties.


## Example

This example turns on the primary value axis for Chart1.


```vb
Charts("Chart1").HasAxis(xlValue, xlPrimary) = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

