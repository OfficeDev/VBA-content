---
title: Curve.End Property (Visio)
keywords: vis_sdr.chm15550575
f1_keywords:
- vis_sdr.chm15550575
ms.prod: visio
api_name:
- Visio.Curve.End
ms.assetid: dce413f4-3c3b-c79f-4dbc-cbe1a8fbcca7
ms.date: 06/08/2017
---


# Curve.End Property (Visio)

Returns the endpoint of a  **Curve** object. Read-only.


## Syntax

 _expression_ . **End**

 _expression_ A variable that represents a **Curve** object.


### Return Value

Double


## Remarks

The  **End** property of a **Curve** object returns the endpoint of a curve. A **Curve** object describes itself in terms of its parameter domain, which is the range [Start(),End()] where End() produces the curve's endpoint.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **End** property to determine the endpoint of a curve.


```vb
 
Sub End_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoPaths As Visio.Paths 
 Dim vsoPath As Visio.Path 
 Dim vsoCurve As Visio.Curve 
 Dim dblStartpoint As Double 
 Dim dblEndpoint As Double 
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
 
 'Display the endpoint of the curve. 
 dblEndpoint = vsoCurve.End 
 Debug.Print "Endpoint= " &; dblEndpoint 
 
 Next intInnerLoopCounter 
 Debug.Print "This path has " &; intInnerLoopCounter - 1 &; " curve object(s)." 
 
 Next intOuterLoopCounter 
 Debug.Print "This shape has " &; intOuterLoopCounter - 1 &; " path object(s)." 
 
End Sub
```


