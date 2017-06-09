---
title: Point.PieSliceLocation Method (PowerPoint)
keywords: vbapp10.chm714011
f1_keywords:
- vbapp10.chm714011
ms.prod: powerpoint
api_name:
- PowerPoint.Point.PieSliceLocation
ms.assetid: 9af5f72b-3626-9f49-09e5-6fdde51f238e
ms.date: 06/08/2017
---


# Point.PieSliceLocation Method (PowerPoint)

Returns the vertical or horizontal position, in points, of a point on a chart item from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

 _expression_. **PieSliceLocation**( **_loc_**, **_Index_** )

 _expression_ A variable that represents a **Point** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _loc_|Required|**[XlPieSliceLocation](http://msdn.microsoft.com/library/d0a2df51-6ab1-8f33-9cdb-29fddc98c058%28Office.15%29.aspx)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional|**[XlPieSliceIndex](http://msdn.microsoft.com/library/04cfc5f3-2a8a-fbd7-e512-4bcd9f524f32%28Office.15%29.aspx)**|Specifies which pie slice position coordinate to return. The default is  **xlOuterCenterPoint**.|

### Return Value

Double


## Remarks

This property applies only to pie chart types.


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

