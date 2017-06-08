---
title: Page.DrawSpline Method (Visio)
keywords: vis_sdr.chm10916230
f1_keywords:
- vis_sdr.chm10916230
ms.prod: visio
api_name:
- Visio.Page.DrawSpline
ms.assetid: a75d7f02-5bfd-f341-ca24-06762e56aca3
ms.date: 06/08/2017
---


# Page.DrawSpline Method (Visio)

Creates a new shape whose path follows a given sequence of points.


## Syntax

 _expression_ . **DrawSpline**( **_xyArray()_** , **_Tolerance_** , **_Flags_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xyArray()_|Required| **Double**|An array of alternating  _x_ and _y_ values that define points in the new shape's path.|
| _Tolerance_|Required| **Double**|How closely the path of the new shape must approximate the given points.|
| _Flags_|Required| **Integer**|Flags that influence how the shape is drawn.|

### Return Value

Shape


## Remarks

The  **DrawSpline** method creates a new shape whose path falls within the given tolerance of the given array of points. To fit the given points exactly, specify a tolerance of zero (0). Typically, the **DrawSpline** method fits spline segments through the points, but it sometimes produces line or circular arc segments in the new shape.

The control points and tolerance are in internal drawing units (inches) with respect to the coordinate space of the page, master, or group in which the new shape is being created. The passed array should be a SAFEARRAY of 8-byte floating point values passed by reference (VT_R8|VT_ARRAY|VT_BYREF). This is how Microsoft Visual Basic passes arrays to Automation objects.

The error from the points to the path of the resulting shape is roughly within tolerance. When the number of points is large, the actual error may sometimes exceed the prescribed tolerance.

The  _Flags_ argument is a bitmask that specifies options for drawing the new shape. Its value should be zero or a combination of one or more of the following values.



|**Constant**|**Value**|
|:-----|:-----|
| **visSplinePeriodic**|1(&;H1)|
| **visSplineDoCircles**|2(&;H2)|
| **visSplineAbrupt**|4(&;H4)|
| **visSpline1D**|8(&;H8)|
If  _Flags_ includes **visSplinePeriodic** and the following conditions are met, the application attempts to draw a periodic spline. Otherwise, Visio draws a non-periodic spline:




- The last point must be a repetition of the first one.
    
- If the flag  **visSplineAbrupt** is included as well, the entire closed path outlined by the points must be free of abrupt changes of direction and curvature.
    


If  _Flags_ includes **visSplineDoCircles** , Microsoft Visio recognizes circular segments in the given array of points and generates circular arcs instead of spline rows for those segments.

If  _Flags_ includes **visSplineAbrupt** , Visio breaks the spline whenever it detects an abrupt change of direction or curvature in the point's trail. An abrupt change of direction is defined by three consecutive points A, B, C in the list, for which the distance between B and the line segment AC is more than twice the tolerance. The application also considers point B to be an abrupt change if one of the segments AB or BC is more than twice as long as the other. At a point where an abrupt change is detected, the application ends the current piece (line, arc, or spline) and starts a fresh one.

If  _Flags_ includes **visSpline1D** and the first and last points in _xyArray()_ don't coincide, the **DrawSpline** method produces a shape that has one-dimensional (1-D) behavior, otherwise, it produces a shape that has two-dimensional (2-D) behavior.

If the first and last points in  _xyArray()_ do coincide, the **DrawSpline** method produces a filled shape.


## Example

The following example shows how to draw a periodic spline through five arbitrary points, requiring that the spline approach within 0.25 (drawing) inches of each point. It allows Visio to start new segments in the path of the new shape at points considered abrupt.


```vb
 
Public Sub DrawSpline_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim intCounter As Integer 
 Dim adblXYPoints(1 To (5 * 2)) As Double 
 
 For intCounter = 1 To 5 
 
 'Set x components (array elements 1,3,5,7,9) to 1,2,3,4,5 
 adblXYPoints((intCounter * 2) - 1) = intCounter 
 
 'Set y components (array elements 2,4,6,8,10) to f(i) 
 adblXYPoints(intCounter * 2) = (intCounter * intCounter) - (7 * intCounter) + 15 
 Next intCounter 
 
 Set vsoShape = ActivePage.DrawSpline(adblXYPoints, 0.25, visSplineAbrupt) 
 
End Sub
```


