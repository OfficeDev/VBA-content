---
title: Selection.FitCurve Method (Visio)
keywords: vis_sdr.chm11116275
f1_keywords:
- vis_sdr.chm11116275
ms.prod: visio
api_name:
- Visio.Selection.FitCurve
ms.assetid: d0f3c799-c15d-cdc8-c0b0-34aeeecec495
ms.date: 06/08/2017
---


# Selection.FitCurve Method (Visio)

Reduces the number of geometry segments in a shape or shapes by replacing them with similar spline, arc, and line segments that approximate the paths of the initial segments. Typically, this reduces the number of segments in the shape.


## Syntax

 _expression_ . **FitCurve**( **_Tolerance_** , **_Flags_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Tolerance_|Required| **Double**|How closely the resulting paths must match the shape's original paths.|
| _Flags_|Required| **Integer**|Flags that influence how the shape is drawn.|

### Return Value

Nothing


## Remarks

The  **FitCurve** method of a **Selection** object optimizes each of the shapes in the selection. It does not combine the selected shapes into a single shape.

The paths resulting from the  **FitCurve** method fall within the given tolerance of the initial paths. Tolerance should be in internal drawing units (inches). To match the initial paths exactly, specify a tolerance of zero (0).

The  _Flags_ argument is a bitmask that specifies options for optimizing the paths. Its value should either be zero or a combination of one or more of the following values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSplinePeriodic**|&;H1|Produce periodic splines if appropriate.|
| **visSplineDoCircles**|&;H2|Recognize circular segments in the shape(s) and generate circular arcs instead of spline rows for those segments.|
| **visSplineAbrupt**|&;H4|Break the resulting splines whenever an abrupt change of direction or curvature in a path is detected.|

