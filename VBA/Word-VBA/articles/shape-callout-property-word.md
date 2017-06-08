---
title: Shape.Callout Property (Word)
keywords: vbawd10.chm161480807
f1_keywords:
- vbawd10.chm161480807
ms.prod: word
api_name:
- Word.Shape.Callout
ms.assetid: 191ba6c5-20e5-458f-b3e3-751a4e566f4a
ms.date: 06/08/2017
---


# Shape.Callout Property (Word)

Returns a  **[CalloutFormat](calloutformat-object-word.md)** object that contains callout formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **Callout**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

This property applies to  **Shape** objects that represent callouts.


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


[Shape Object](shape-object-word.md)

