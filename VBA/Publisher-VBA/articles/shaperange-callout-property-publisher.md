---
title: ShapeRange.Callout Property (Publisher)
keywords: vbapb10.chm2293811
f1_keywords:
- vbapb10.chm2293811
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Callout
ms.assetid: 25b9b444-6cbf-085a-df7f-8899e8e55057
ms.date: 06/08/2017
---


# ShapeRange.Callout Property (Publisher)

Returns a  **[CalloutFormat](calloutformat-object-publisher.md)** object representing the formatting of a line callout.


## Syntax

 _expression_. **Callout**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example adds an oval to the active publication and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Sub NewShapeItem() 
 
 Dim shpNew As Shapes 
 
 Set shpNew = Application.ActiveDocument.MasterPages(1).Shapes 
 With shpNew 
 .AddShape Type:=msoShapeOval, Left:=180, _ 
 Top:=200, Width:=280, Height:=130 
 With .AddCallout(Type:=msoCalloutTwo, Left:=420, _ 
 Top:=170, Width:=170, Height:=40) 
 .TextFrame.TextRange = "Big Oval" 
 With .Callout 
 .Accent = msoTrue 
 .Border = msoFalse 
 End With 
 End With 
 End With 
 
End Sub
```


