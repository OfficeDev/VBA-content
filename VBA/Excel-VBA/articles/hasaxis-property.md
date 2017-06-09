---
title: HasAxis Property
keywords: vbagr10.chm65588
f1_keywords:
- vbagr10.chm65588
ms.prod: excel
api_name:
- Excel.HasAxis
ms.assetid: 2de3c3a1-7b9c-a4d9-40cb-906fd5d6f4cb
ms.date: 06/08/2017
---


# HasAxis Property

Returns or sets which axes exist on the chart. Read/write Variant.

 _expression_. **HasAxis**( **_Index1_**,  **_Index2_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index1** Optional **XlAxisType**. The type of axis.


|XlAxisType can be one of these XlAxisType constants.|
| **xlCategory**|
| **xlValue** **xlSeriesAxis**. Series axes apply only to 3-D charts.|
 **Index2** Optional **XlAxisGroup**. The axis priority.


|XlAxisGroup can be one of these XlAxisGroup constants.|
| **xlPrimary**|
| **xlSecondary**3-D charts have only one set of axes.|

## Remarks

Microsoft Graph may create or delete axes if you change the chart type or change the  **[AxisGroup](axisgroup-property.md)** property.


## Example

This example turns on the primary value axis.


```vb
myChart.HasAxis(xlValue, xlPrimary) = True
```


