---
title: Point.PieSliceLocation Method (Word)
keywords: vbawd10.chm262146656
f1_keywords:
- vbawd10.chm262146656
ms.prod: word
api_name:
- Word.Point.PieSliceLocation
ms.assetid: 85687cf7-b9a8-a51d-886c-c45092cbd929
ms.date: 06/08/2017
---


# Point.PieSliceLocation Method (Word)

Returns the vertical or horizontal position of a point on a chart item, in points, from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

 _expression_ . **PieSliceLocation**( **_loc_** , **_Index_** )

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _loc_|Required| **[XlPieSliceLocation](xlpieslicelocation-enumeration-word.md)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional| **[XlPieSliceIndex](xlpiesliceindex-enumeration-word.md)**|Specifies which pie slice position coordinate to return. The default value is  **xlOuterCenterPoint** .|

### Return Value

Double


## Remarks

This property only applies to pie chart types.


## See also


#### Concepts


[Point Object](point-object-word.md)

