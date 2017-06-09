---
title: Point.PieSliceLocation Method (Excel)
keywords: vbaxl10.chm576109
f1_keywords:
- vbaxl10.chm576109
ms.prod: excel
api_name:
- Excel.Point.PieSliceLocation
ms.assetid: 90a318d4-0ad2-d326-c26b-3c965b1ffe43
ms.date: 06/08/2017
---


# Point.PieSliceLocation Method (Excel)

Returns the vertical or horizontal position of a point on a chart item, in points, from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

 _expression_ . **PieSliceLocation**( **_loc_** , **_Index_** )

 _expression_ A variable that represents a **[Point](point-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _loc_|Required| **[XlPieSliceLocation](xlpieslicelocation-enumeration-excel.md)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional| **[XlPieSliceIndex](xlpiesliceindex-enumeration-excel.md)**|Specifies which pie slice position coordinate to return. The default value is  **xlOuterCenterPoint** .|

### Return Value

Double


## Remarks

This property only applies to pie and doughnut chart types.


## See also


#### Concepts


[Point Object](point-object-excel.md)

