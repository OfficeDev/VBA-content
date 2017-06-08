---
title: TextRange.Text Property (Publisher)
keywords: vbapb10.chm5308416
f1_keywords:
- vbapb10.chm5308416
ms.prod: publisher
api_name:
- Publisher.TextRange.Text
ms.assetid: 13584812-307a-c32b-ca8f-27869728b64e
ms.date: 06/08/2017
---


# TextRange.Text Property (Publisher)

Returns or sets a  **String** that represents the text in a text range or WordArt shape. Read/write.


## Syntax

 _expression_. **Text**

 _expression_A variable that represents a  **TextRange** object.


## Example

The following example adds a rectangle to the active publication and adds text to it.


```vb
Sub AddTextToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
 Left:=72, Top:=72, Width:=250, Height:=140) 
 .TextFrame.TextRange.Text = "Here is some test text" 
 End With 
End Sub
```


