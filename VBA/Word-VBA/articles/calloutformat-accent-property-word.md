---
title: CalloutFormat.Accent Property (Word)
keywords: vbawd10.chm163905636
f1_keywords:
- vbawd10.chm163905636
ms.prod: word
api_name:
- Word.CalloutFormat.Accent
ms.assetid: 7c6d7e02-5117-36ab-1d61-72ef9c4b0fd3
ms.date: 06/08/2017
---


# CalloutFormat.Accent Property (Word)

 **True** if a vertical accent bar separates the callout text from the callout line. Read/write **MsoTriState** .


## Syntax

 _expression_ . **Accent**

 _expression_ A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


## Example

This example adds an oval to the active document and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Dim shapeCallout As Shape 
 
With ActiveDocument.Shapes 
 .AddShape msoShapeOval, 180, 200, 280, 130 
 Set shapeCallout = .AddCallout(msoCalloutTwo, 420, 170, 170, 40) 
 
 With shapeCallout 
 .TextFrame.TextRange.Text = "My oval" 
 .Callout.Accent = msoTrue 
 .Callout.Border = msoFalse 
 End With 
End With 

```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

