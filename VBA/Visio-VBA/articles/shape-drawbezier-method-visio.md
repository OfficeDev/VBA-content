---
title: Shape.DrawBezier Method (Visio)
keywords: vis_sdr.chm11216195
f1_keywords:
- vis_sdr.chm11216195
ms.prod: visio
api_name:
- Visio.Shape.DrawBezier
ms.assetid: d38b56a5-2366-e418-206f-db39bd8e2c82
ms.date: 06/08/2017
---


# Shape.DrawBezier Method (Visio)

Creates a shape whose path is defined by the supplied sequence of Bezier control points.


## Syntax

 _expression_ . **DrawBezier**( **_xyArray()_** , **_degree_** , **_Flags_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xyArray()_|Required| **Double**|An array of alternating  _x_ and _y_ values that define the Bezier control points for the new shape.|
| _degree_|Required| **Integer**|The degree of the Bezier curve.|
| _Flags_|Required| **Integer**|Flags that influence how the shape is drawn.|

### Return Value

Shape


## Remarks

The  _xyArray()_ and _degree_ parameters must meet the following conditions:

1 <=  _degree_ <= 9

The number of points must be _ k * degree_ + 1, where _k_ is a positive integer. If the first point is called _p0_ , for any integer _m_ between 1 and _k_ , _p(m * degree)_ is assumed to be the last control point of a Bezier segment, as well as the first control point of the next.

The result is a composite curve that consists of  _k_ Bezier segments. The input points from _xyArray()_ define the curve's control points. If you want a smooth curve, make sure the points _p(n - 1)_ , _pn_ , and _p(n + 1)_ are co-linear whenever _n = m * degree_ with an integer _m_ . The composite Bezier curve is represented in the application as a B-spline with integer _knots_ of _multiplicity = degree_ .

The control points should be in internal drawing units (inches) with respect to the coordinate space of the page, master, or group where the shape is being drawn. The passed array should be a SAFEARRAY of 8-byte floating point values passed by reference (VT_R8|VT_ARRAY|VT_BYREF). This is how Microsoft Visual Basic passes arrays to Automation objects.

The  _Flags_ argument is a bitmask that specifies options for drawing the new shape. Its value should be zero (0) or **visSpline1D** (8).

If  _Flags_ is **visSpline1D** and the first and last points in _xyArray()_ don't coincide, the **DrawBezier** method produces a shape with one-dimensional (1-D) behavior; otherwise, it produces a shape with two-dimensional (2-D) behavior.

If the first and last points in  _xyArray()_ do coincide, the **DrawBezier** method produces a filled shape.


## Example

The following example shows how to draw a Bezier curve through five arbitrary points on the active page.


```vb
 
Public Sub DrawBezier_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim intCounter As Integer 
 Dim adblXYPoints(1 To (5 * 2)) As Double 
 
 For intCounter = 1 To 5 
 
 'Set x-coordinates (array elements 1,3,5,7,9) to 1,2,3,4,5 
 adblXYPoints((intCounter * 2) - 1) = intCounter 
 
 'Set y-coordinates (array elements 2,4,6,8,10) to f(intCounter) 
 adblXYPoints(intCounter * 2) = (intCounter * intCounter) - (7 * intCounter) + 15 
 
 Next intCounter 
 
 Set vsoShape = ActivePage.DrawBezier(adblXYPoints, 2, visSpline1D) 
 
End Sub
```


