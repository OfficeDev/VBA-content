---
title: ShapeRange.Callout Property (Word)
keywords: vbawd10.chm162857063
f1_keywords:
- vbawd10.chm162857063
ms.prod: word
api_name:
- Word.ShapeRange.Callout
ms.assetid: 87cc8811-497d-17b9-4483-682cdd1fbce3
ms.date: 06/08/2017
---


# ShapeRange.Callout Property (Word)

Returns a  **[CalloutFormat](calloutformat-object-word.md)** object that contains callout formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **Callout**

 _expression_ A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

This property applies to  **ShapeRange** objects that represent callouts.


## Example

This example adds to myDocument an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes 
 .AddShape msoShapeOval, 180, 200, 280, 130 
 With .AddCallout(msoCalloutTwo, 420, 170, 170, 40) 
 .TextFrame.TextRange.Text = "My oval" 
 With .Callout 
 .Accent = True 
 .Border = False 
 End With 
 End With 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

