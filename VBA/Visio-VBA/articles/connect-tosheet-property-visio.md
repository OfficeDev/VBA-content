---
title: Connect.ToSheet Property (Visio)
keywords: vis_sdr.chm10314585
f1_keywords:
- vis_sdr.chm10314585
ms.prod: visio
api_name:
- Visio.Connect.ToSheet
ms.assetid: 449993f6-dd44-cebf-8d2d-343e0202b166
ms.date: 06/08/2017
---


# Connect.ToSheet Property (Visio)

Returns the shape to which one or more connections are made. Read-only.


## Syntax

 _expression_ . **ToSheet**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Shape


## Remarks

The  **ToSheet** property for a **Connect** object always returns the shape to which the connection is made.

The  **Connects** collection represents several connections. If every connection represented by the collection is made to the same shape, the **ToSheet** property returns that shape. Otherwise, it returns **Nothing** and does not raise an exception.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ToSheet** property to find the shape a **Connect** object originates from in a Microsoft Visio drawing. The example displays the connection information in the Immediate window.

This example assumes there is an active document that contains at least two connected shapes. For best results, connect two shapes from the  **Organization Chart Shapes** stencil.




```vb
 
Public Sub ToSheet_Example() 
 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoShape As Visio.Shape 
 Dim vsoConnectTo As Visio.Shape 
 Dim vsoConnects As Visio.Connects 
 Dim vsoConnect As Visio.Connect 
 Dim intCurrentShapeIndex As Integer 
 Dim intCounter As Integer 
 Set vsoShapes = ActivePage.Shapes 
 
 'For each shape on the page, get its connections. 
 For intCurrentShapeIndex = 1 To vsoShapes.Count 
 
 Set vsoShape = vsoShapes(intCurrentShapeIndex) 
 Set vsoConnects = vsoShape.Connects 
 
 'For each connection, get the shape it connects to. 
 For intCounter = 1 To vsoConnects.Count 
 
 Set vsoConnect = vsoConnects(intCounter) 
 Set vsoConnectTo = vsoConnect.ToSheet 
 
 'Print the name of the shape the 
 'Connect object connects to. 
 Debug.Print vsoConnectTo.Name 
 
 Next intCounter 
 
 Next intCurrentShapeIndex 
 
End Sub
```


