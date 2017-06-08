---
title: Curve.PointAndDerivatives Method (Visio)
keywords: vis_sdr.chm15516440
f1_keywords:
- vis_sdr.chm15516440
ms.prod: visio
api_name:
- Visio.Curve.PointAndDerivatives
ms.assetid: 2df3753b-f0f5-37ff-75d9-f63d6fc491dc
ms.date: 06/08/2017
---


# Curve.PointAndDerivatives Method (Visio)

Returns a point and its derivatives at a position along a curve's path.


## Syntax

 _expression_ . **PointAndDerivatives**( **_t_** , **_n_** , **_x_** , **_y_** , **_dxdt_** , **_dydt_** , **_ddxdt_** , **_ddydt_** )

 _expression_ A variable that represents a **Curve** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _t_|Required| **Double**|The value in the curve's parameter domain to evaluate.|
| _n_|Required| **Integer**|0: get point; 1: point and 1st derivative; 2: point plus first and second derivative.|
| _x_|Required| **Double**|Returns  _x_ value of curve at _t._|
| _y_|Required| **Double**|Returns  _y_ value of curve at _t_.|
| _dxdt_|Required| **Double**|Returns first derivative ( _dx/dt_) at  _t_ if _n_ > 0.|
| _dydt_|Required| **Double**|Returns first derivative ( _dy/dt_) at  _t_ if _n_> 0.|
| _ddxdt_|Required| **Double**|Returns second derivative ( _ddx/dt_) at  _t_ if _n_> 1.|
| _ddydt_|Required| **Double**|Returns second derivative ( _ddy/dt_) at  _t_ if _n_> 1.|

### Return Value

Nothing


## Remarks

Use the  **PointAndDerivatives** method of the **Curve** object to obtain the coordinates of a point within the curve's parameter domain and its first and second derivatives.

A  **Curve** object is described in terms of its parameter domain, which is the range [Start(),End()]. The **PointAndDerivatives** method can be used to extrapolate the curve's path outside [Start(),End()].


## Example

This Microsoft Visual Basic for Applications (VBA) macro draws an oval on the document's active page and then retrieves it and iterates through its  **Paths** collection and each **Path** object to display the coordinates of various points along the curve. Because the shape drawn is an oval, it contains only one path and only one **Curve** object.


```vb
 
Sub PointAndDerivatives_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoPaths As Visio.Paths 
 Dim vsoPath As Visio.Path 
 Dim vsoCurve As Visio.Curve 
 Dim dblStartpoint As Double 
 Dim dblXCoordinate As Double 
 Dim dblYCoordinate As Double 
 Dim dblFirstDerivativeX As Double 
 Dim dblFirstDerivativeY As Double 
 Dim dblSecondDerivativeX As Double 
 Dim dblSecondDerivativeY As Double 
 Dim intOuterLoopCounter As Integer 
 Dim intInnerLoopCounter As Integer 
 
 'Get the Paths collection for this shape. 
 Set vsoPaths = ActivePage.DrawOval(1, 1, 4, 4).Paths 
 
 'Iterate through the Path objects in the Paths collection. 
 For intOuterLoopCounter = 1 To vsoPaths.Count 
 Set vsoPath = vsoPaths.Item(intOuterLoopCounter) 
 Debug.Print "Path object " &; intOuterLoopCounter 
 
 'Iterate through the curves in a Path object. 
 For intInnerLoopCounter = 1 To vsoPath.Count 
 
 Set vsoCurve = vsoPath(intInnerLoopCounter) 
 Debug.Print "Curve number " &; intInnerLoopCounter 
 
 'Display the start point of the curve. 
 dblStartpoint = vsoCurve.Start 
 Debug.Print "Startpoint= " &; dblStartpoint 
 
 'Use the PointAndDerivatives method to obtain 
 'a point and the first derivative at that point. 
 vsoCurve.PointAndDerivatives (dblStartpoint - 1), 1, _ 
 dblXCoordinate, dblYCoordinate, dblFirstDerivativeX, dblFirstDerivativeY, dblSecondDerivativeX, dblSecondDerivativeY 
 Debug.Print "PointAndDerivative= " &; dblXCoordinate, dblYCoordinate, dblFirstDerivativeX, dblFirstDerivativeY 
 
 Next intInnerLoopCounter 
 Debug.Print "This path has " &; intInnerLoopCounter - 1 &; " curve object(s)." 
 
 Next intOuterLoopCounter 
 Debug.Print "This shape has " &; intOuterLoopCounter - 1 &; " path object(s)." 
 
End Sub
```


