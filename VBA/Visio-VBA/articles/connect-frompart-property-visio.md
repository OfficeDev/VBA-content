---
title: Connect.FromPart Property (Visio)
keywords: vis_sdr.chm10313585
f1_keywords:
- vis_sdr.chm10313585
ms.prod: visio
api_name:
- Visio.Connect.FromPart
ms.assetid: 3ef8eaf8-b405-057d-6afd-ccfa16dfab62
ms.date: 06/08/2017
---


# Connect.FromPart Property (Visio)

Returns the part of a shape from which a connection originates. Read-only.


## Syntax

 _expression_ . **FromPart**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Integer


## Remarks

The following constants declared by the Microsoft Visio type library show return values for the  **FromPart** property.



|**Constant**|**Value**|
|:-----|:-----|
| **visConnectFromError**|-1|
| **visFromNone**|0|
| **visLeftEdge**|1|
| **visCenterEdge**|2|
| **visRightEdge**|3|
| **visBottomEdge**|4|
| **visMiddleEdge**|5|
| **visTopEdge**|6|
| **visBeginX**|7|
| **visBeginY**|8|
| **visBegin**|9|
| **visEndX**|10|
| **visEndY**|11|
| **visEnd**|12|
| **visFromAngle**|13|
| **visFromPin**|14|
| **visControlPoint**|100 + zero-based row index (for example,  **visControlPoint** = 100 if the control point is in row 0; **visControlPoint** = 101 if the control point is in row 1)|

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to extract connection information from a Visio drawing. The example displays the connection information in the Immediate window.



This example assumes there is an active document that contains at least two connected shapes.




```vb
 
Public Sub FromPart_Example() 
 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoShape As Visio.Shape 
 Dim vsoConnectFrom As Visio.Shape 
 Dim intFromData As Integer 
 Dim strFrom As String 
 Dim vsoConnects As Visio.Connects 
 Dim vsoConnect As Visio.Connect 
 Dim intCurrentShapeIndex As Integer 
 Dim intCounter As Integer 
 Set vsoShapes = ActivePage.Shapes 
 
 'For each shape on the page, get its connections. 
 For intCurrentShapeIndex = 1 To vsoShapes.Count 
 Set vsoShape = vsoShapes(intCurrentShapeIndex) 
 Set vsoConnects = vsoShape.Connects 
 
 'For each connection, get the shape it originates from 
 'and the part of the shape it originates from, 
 'and print that information in the Immediate window. 
 For intCounter = 1 To vsoConnects.Count 
 Set vsoConnect = vsoConnects(intCounter) 
 Set vsoConnectFrom = vsoConnect.FromSheet 
 intFromData = vsoConnect.FromPart 
 
 'FromPart property values 
 If intFromData = visConnectError Then 
 strFrom = "error" 
 ElseIf intFromData = visNone Then 
 strFrom = "none" 
 ElseIf intFromData = visLeftEdge Then 
 strFrom = "left" 
 ElseIf intFromData = visCenterEdge Then 
 strFrom = "center" 
 ElseIf intFromData = visRightEdge Then 
 strFrom = "right" 
 ElseIf intFromData = visBottomEdge Then 
 strFrom = "bottom" 
 ElseIf intFromData = visMiddleEdge Then 
 strFrom = "middle" 
 ElseIf intFromData = visTopEdge Then 
 strFrom = "top" 
 ElseIf intFromData = visBeginX Then 
 strFrom = "beginX" 
 ElseIf intFromData = visBeginY Then 
 strFrom = "beginY" 
 ElseIf intFromData = visBegin Then 
 strFrom = "begin" 
 ElseIf intFromData = visEndX Then 
 strFrom = "endX" 
 ElseIf intFromData = visEndY Then 
 strFrom = "endY" 
 ElseIf intFromData = visEnd Then 
 strFrom = "end" 
 ElseIf intFromData >= visControlPoint Then 
 strFrom = "controlPt_" &; _ 
 Str(intFromData - visControlPoint + 1) 
 Else 
 strFrom = "???" 
 End If 
 
 Debug.Print vsoConnectFrom.Name &; " " &; strFrom 
 
 Next intCounter 
 
 Next intCurrentShapeIndex 
 
End Sub
```


