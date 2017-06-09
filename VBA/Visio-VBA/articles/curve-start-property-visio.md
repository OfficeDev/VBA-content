---
title: Curve.Start Property (Visio)
keywords: vis_sdr.chm15514410
f1_keywords:
- vis_sdr.chm15514410
ms.prod: visio
api_name:
- Visio.Curve.Start
ms.assetid: ac5e56e8-dad2-c150-02e4-f5d7dafd20ff
ms.date: 06/08/2017
---


# Curve.Start Property (Visio)

Returns the start of a  **Curve** object's parameter domain. Read-only.


## Syntax

 _expression_ . **Start**

 _expression_ A variable that represents a **Curve** object.


### Return Value

Double


## Remarks

The  **Start** property of a **Curve** object returns the value of the starting point in the curve's parameter domain. A **Curve** object describes itself in terms of its parameter domain, which is the range [Start(),End()], where Start() produces the curve's starting point. Note that the **Start** value is not a coordinate pair. Rather, it represents the relative position along the curve of the starting point. For a line, for example, the value of **Start** typically is 0, the value of **End** is 1, and you can use the **Point** method of the **Curve** object to determine the coordinates of any point along the curve by determining the relative location of the point between the start and endpoints.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Start** property to display the value of the starting point of a curve. It uses the **Point** method to find the midpoint of the curve.


```vb
 
Sub Start_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoPaths As Visio.Paths 
 Dim vsoPath As Visio.Path 
 Dim vsoCurve As Visio.Curve 
 Dim dblStartpoint As Double 
 Dim dblEndpoint As Double 
 Dim dblX As Double 
 Dim dblY As Double 
 Dim intOuterLoopCounter As Integer 
 Dim intInnerLoopCounter As Integer 
 
 'Draw a shape and get its Paths collection. 
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
 Debug.Print "Startpoint = " &; dblStartpoint 
 
 'Display the endpoint of the curve. 
 dblEndpoint = vsoCurve.End 
 Debug.Print "Endpoint = " &; dblEndpoint 
 
 'Find the midpoint of the curve. 
 vsoCurve.Point ((dblEndpoint - dblStartpoint) / 2), dblX, dblY 
 Debug.Print "Midpoint: x = " &; dblx; ", y = " &; dblY 
 
 Next intInnerLoopCounter 
 Debug.Print "This path has " &; intInnerLoopCounter - 1 &; " curve object(s)." 
 
 Next intOuterLoopCounter 
 Debug.Print "This shape has " &; intOuterLoopCounter - 1 &; " path object(s)." 
 
End Sub
```


