---
title: Path.Points Method (Visio)
keywords: vis_sdr.chm15414090
f1_keywords:
- vis_sdr.chm15414090
ms.prod: visio
api_name:
- Visio.Path.Points
ms.assetid: 3c14acdb-33ad-9bd9-96d9-320bd53fa5c9
ms.date: 06/08/2017
---


# Path.Points Method (Visio)

Returns an array of points that defines a polyline that approximates a  **Path** or **Curve** object within a given tolerance.


## Syntax

 _expression_ . **Points**( **_Tolerance_** , **_xyArray()_** )

 _expression_ A variable that represents a **Path** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Tolerance_|Required| **Double**|Specifies how close the returned array of points must approximate the true path.|
| _xyArray()_|Required| **Double**|Out parameter. Returns an array of alternating  _x_ and _y_ values specifying points along a path's or curve's stroke.|

### Return Value

Nothing


## Remarks

Use the  **Points** method of the **Path** or **Curve** object to obtain an array of _x,y_ coordinates specifying points along the path or curve within a given tolerance. The tolerance and returned _x,y_ values are expressed in internal drawing units (inches).

If you used the  **Paths** property of a **Shapes** object to obtain the **Path** or **Curve** object being queried, the coordinates are expressed in the parent's coordinate system. If you used the **PathsLocal** property of a **Shape** object to obtain the **Path** or **Curve** object, the coordinates are expressed in the local coordinate system.

If Microsoft Visio is unable to achieve the requested tolerance, Visio approximates the points as close to the requested tolerance as possible. Generally speaking, the lower the tolerance, the more points Visio returns. Visio doesn't accept a tolerance of zero (0).

The array returned includes both the starting and ending points of the path or curve even if it is closed.


## Example

This Microsoft Visual Basic for Applications (VBA) macro places a shape on the page, retrieves its  **Paths** collection, and then uses the **Points** method of the **Path** object to return an array of points that defines a polyline approximating the **Path** object.


```vb
 
Public Sub Points_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim adblXYPoints() As Double 
 Dim strPointsList As String 
 Dim intOuterLoopCounter As Integer 
 Dim intInnerLoopCounter As Integer 
 
 Set vsoShape = ActivePage.DrawOval(1, 1, 4, 4) 
 
 For intOuterLoopCounter = 1 To vsoShape.Paths.Count 
 
 vsoShape.Paths(intOuterLoopCounter).Points 0.1, adblXYPoints 
 For intInnerLoopCounter = LBound(adblXYPoints) To UBound(adblXYPoints) 
 strPointsList = strPointsList &; adblXYPoints(intInnerLoopCounter) &; Chr(10) 
 Next intInnerLoopCounter 
 
 Next intOuterLoopCounter 
 
 Debug.Print strPointsList 
 
End Sub
```


