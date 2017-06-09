---
title: Shape.DistanceFromPoint Property (Visio)
keywords: vis_sdr.chm11213425
f1_keywords:
- vis_sdr.chm11213425
ms.prod: visio
api_name:
- Visio.Shape.DistanceFromPoint
ms.assetid: 262b5814-3b86-c3eb-9526-96ec73836ad6
ms.date: 06/08/2017
---


# Shape.DistanceFromPoint Property (Visio)

Returns the distance from a shape to a point. Read-only.


## Syntax

 _expression_ . **DistanceFromPoint**( **_x_** , **_y_** , **_Flags_** , **_[pvPathIndex]_** , **_[pvCurveIndex]_** , **_[pvt]_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Double**|An x-coordinate.|
| _y_|Required| **Double**|A y-coordinate.|
| _Flags_|Required| **Integer**|Flags that influence the type of entries returned in results.|
| _pvPathIndex_|Optional| **Variant**|Identifies the point on the shape in conjunction with  _pvCurveIndex_ and _pvt_.|
| _pvCurveIndex_|Optional| **Variant**|Identifies the point on the shape in conjunction with  _pvPathIndex_ and _pvt_.|
| _pvt_|Optional| **Variant**|Identifies the point on the shape in conjunction with  _pvPathIndex_ and _pvCurveIndex_.|

### Return Value

Double


## Remarks

The ( _x,y_) point is expressed in internal drawing units (inches in the drawing) with respect to the coordinate space defined by the sheet immediately containing ThisShape.

The  _pvPathIndex_,  _pvCurveIndex_, and  _pvt_ arguments optionally return values that identify the point the returned distance is measured from. Call that point ( _xOnThis,yOnThis_). It lies along the  _c_'th curve of ThisShape's  _p_'th path and can be determined by:




```
ThisShape.Paths(*pvPathIndex).Item(*pvCurveIndex).Point(*pvt,&;xOnThis ,&;yOnthis)
```

You can use the  **PointAndDerivatives** method instead of the **Point** method if you want to find the first and second derivatives at position _t_ along the curve.

If  _pvPathIndex_ or _pvCurveIndex_ is not **Null** , an **Integer** (type VT_I4) is returned. If _pvt_ isn't **Null** , **DistanceFromPoint** returns a **Double** (type VT_R8).

The  **DistanceFromPoint** property considers guides to have extent and considers a shape's filled areas and paths.

The  _Flags_ argument can be any combination of the values of the constants defined in the following table. These constants are also defined in **VisSpatialRelationFlags** in the Microsoft Visio type library.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSpatialIncludeDataGraphics**|&;H40|Includes data graphic callout shapes and their sub-shapes. By default, data graphic callout shapes and their subshapes are not included. If the parent shape is itself a data graphic callout, searches are made between the parent shape's geometry and non-callout shapes, unless this flag is set.|
| **visSpatialIncludeHidden**|&;H10 |Consider hidden Geometry sections. By default, hidden Geometry sections do not influence the result. |
| **visSpatialIgnoreVisible**|&;H20 |Do not consider visible Geometry sections. By default, visible Geometry sections influence the result. |
Use the NoShow cell to determine whether a Geometry section is hidden or visible. Hidden Geometry sections have a value of TRUE and visible Geometry sections have a value of FALSE in the NoShow cell.

If the parent object has no geometry, or if  _Flags_ excludes consideration of all geometry, the **DistanceFromPoint** property returns a large number (1E+30) which should be interpreted as infinite.

The  **DistanceFromPoint** property does not consider the width of a shape's line, shadows, line ends, control points, or connection points when computing its result.


