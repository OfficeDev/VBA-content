---
title: Curve.Point Method (Visio)
keywords: vis_sdr.chm15516435
f1_keywords:
- vis_sdr.chm15516435
ms.prod: visio
api_name:
- Visio.Curve.Point
ms.assetid: 48fcad31-a655-f68c-10fd-127fea45f95d
ms.date: 06/08/2017
---


# Curve.Point Method (Visio)

Returns a point at a position along a curve.


## Syntax

 _expression_ . **Point**( **_t_** , **_x_** , **_y_** )

 _expression_ A variable that represents a **Curve** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _t_|Required| **Double**|The value in the curve's parameter domain to evaluate.|
| _x_|Required| **Double**|Returns  _x_ value of curve at _t_.|
| _y_|Required| **Double**|Returns  _y_ value of curve at _t_.|

### Return Value

Nothing


## Remarks

A  **Curve** object is described in terms of its parameter domain, which is the range [Start(),End()]. The **Point** method of a **Curve** object returns the _x,y_ coordinates at position _t_, which is any position along the curve's path. The  **Point** method can be used to extrapolate the curve's path outside [Start(),End()].


## Example

This Microsoft Visual Basic for Applications (VBA) macro draws a circle (a special case of an oval) on the document's active page. Then it iterates through the  **Paths** collection of the circle and each **Path** object to display the coordinates of various points along the curve. Because the shape drawn is a circle, it is a **Curve** object that has only one path.


```vb
 
Sub Point_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoPaths As Visio.Paths 
 Dim vsoPath As Visio.Path 
 Dim vsoCurve As Visio.Curve 
 Dim dblEndpoint As Double 
 Dim dblXCoordinate As Double 
 Dim dblYCoordinate As Double 
 Dim intOuterLoopCounter As Integer 
 Dim intInnerLoopCounter As Integer 
 
 'Get the Paths collection for this shape. 
 Set vsoPaths = ActivePage.DrawOval(1, 1, 4, 4).Paths 
 
 'Iterate through the Path objects in the Paths collection. 
 For intOuterLoopCounter = 1 To vsoPaths.Count 
 Set vsoPath = vsoPaths.Item(intOuterLoopCounter) 
 Debug.Print "Path object " &; intOuterLoopCounter 
 
 'Iterate through the curves in the Path object. 
 For intInnerLoopCounter = 1 To vsoPath.Count 
 
 Set vsoCurve = vsoPath(intInnerLoopCounter) 
 Debug.Print "Curve number " &; intInnerLoopCounter 
 
 'Display the endpoint of the curve 
 dblEndpoint = vsoCurve.End 
 Debug.Print "Endpoint= " &; dblEndpoint 
 
 'Use the Point method to determine the 
 'coordinates of an arbitrary point on the curve 
 vsoCurve.Point (dblEndpoint/2), dblXCoordinate, dblYCoordinate 
 Debug.Print "Point= " &; dblXCoordinate, dblYCoordinate 
 
 Next intInnerLoopCounter 
 Debug.Print "This path has " &; intInnerLoopCounter - 1 &; " curve object(s)." 
 
 Next intOuterLoopCounter 
 Debug.Print "This shape has " &; intOuterLoopCounter - 1 &; " path object(s)." 
 
End Sub
```


