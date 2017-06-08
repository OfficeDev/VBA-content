---
title: Connect.ToPart Property (Visio)
keywords: vis_sdr.chm10314575
f1_keywords:
- vis_sdr.chm10314575
ms.prod: visio
api_name:
- Visio.Connect.ToPart
ms.assetid: 37044045-f911-872e-4f72-68fa265fb6f8
ms.date: 06/08/2017
---


# Connect.ToPart Property (Visio)

Returns the part of a shape to which a connection is made. Read-only.


## Syntax

 _expression_ . **ToPart**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Integer


## Remarks

The  **ToPart** property identifies the part of a shape to which another shape is glued, such as its begin point or endpoint, one of its edges, or a connection point. The following constants declared by the Visio type library in member **VisToParts** show possible return values for the **ToPart** property.



|**Constant**|**Value**|
|:-----|:-----|
| **visConnectToError**|-1|
| **visToNone**|0|
| **visGuideX**|1|
| **visGuideY**|2|
| **visWholeShape**|3|
| **visGuideIntersect**|4|
| **visToAngle**|7|
| **visConnectionPoint**|100 + row index of connection point|

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to extract connection information from a Microsoft Visio drawing. The example displays the connection information in the Immediate window.



This example assumes there is an active document that contains at least two connected shapes.




```vb
 
Public Sub ToPart_Example() 
 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoShape As Visio.Shape 
 Dim vsoConnectTo As Visio.Shape 
 Dim intToData As Integer 
 Dim strTo As String 
 Dim vsoConnects As Visio.Connects 
 Dim vsoConnect As Visio.Connect 
 Dim intCurrentShapeID As Integer 
 Dim intCounter As Integer 
 
 Set vsoShapes = ActivePage.Shapes 
 
 'For each shape on the page, get its connections. 
 For intCurrentShapeID = 1 To vsoShapes.Count 
 
 Set vsoShape = vsoShapes(intCurrentShapeID) 
 Set vsoConnects = vsoShape.Connects 
 
 'For each connection, get the shape it connects to 
 'and the part of the shape it connects to, 
 'and print that information in the Immediate window. 
 For intCounter = 1 To vsoConnects.Count 
 
 Set vsoConnect = vsoConnects(intCounter) 
 Set vsoConnectTo = vsoConnect.ToSheet 
 intToData = vsoConnect.ToPart 
 
 If intToData = visConnectError Then 
 strTo = "error" 
 ElseIf intToData = visNone Then 
 strTo = "none" 
 ElseIf intToData = visGuideX Then 
 strTo = "guideX" 
 ElseIf intToData = visGuideY Then 
 strTo = "guideY" 
 ElseIf intToData = visWholeShape Then 
 strTo = "dynamic glue" 
 ElseIf intToData >= visConnectionPoint Then 
 strTo = "connection point " &; _ 
 CStr(intToData - visConnectionPoint + 1) 
 Else 
 strTo = "???" 
 End If 
 
 'Print the name and part of the shape the 
 'Connect object connects to. 
 Debug.Print "To "; vsoConnectTo.Name &; " " &; strTo &; "." 
 
 Next intCounter 
 
 Next intCurrentShapeID 
 
End Sub
```


