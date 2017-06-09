---
title: Shape.Paths Property (Visio)
keywords: vis_sdr.chm11214055
f1_keywords:
- vis_sdr.chm11214055
ms.prod: visio
api_name:
- Visio.Shape.Paths
ms.assetid: 8a179059-7cab-728a-c7b8-a4d8b31476ee
ms.date: 06/08/2017
---


# Shape.Paths Property (Visio)

Returns a  **Paths** collection that reports the coordinates of a shape's paths in the coordinate system of the shape's parent. Read-only.


## Syntax

 _expression_ . **Paths**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Paths


## Example

This Microsoft Visual Basic for Applications (VBA) macro places a shape on the page, retrieves its  **Paths** collection, and then uses the **Points** property of the **Path** object to return an array of points that defines a polyline approximating the **Path** object.


```vb
 
Public Sub Paths_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim adblXYPoints() As Double 
 Dim strPointsList As String 
 Dim intOuterLoopCounter As Integer 
 Dim intInnerLoopCounter As Integer 
 
 Set vsoShape = ActivePage.DrawOval(1, 1, 4, 4) 
 
 For intOuterLoopCounter = 1 To vsoShape.Paths.Count 
 
 vsoShape.Paths(intOuterLoopCounter).Points 1#, adblXYPoints 
 For intInnerLoopCounter = LBound(adblXYPoints) To UBound(adblXYPoints) 
 strPointsList = strPointsList &; adblXYPoints(intInnerLoopCounter) &; Chr(10) 
 Next intInnerLoopCounter 
 
 Next intOuterLoopCounter 
 Debug.Print strPointsList 
 
End Sub
```


