---
title: ShapeRange.TextEffect Property (Publisher)
keywords: vbapb10.chm2293833
f1_keywords:
- vbapb10.chm2293833
ms.prod: publisher
api_name:
- Publisher.ShapeRange.TextEffect
ms.assetid: 7bc822f2-4754-685d-fdd3-7479b5a3ac52
ms.date: 06/08/2017
---


# ShapeRange.TextEffect Property (Publisher)

Returns a  **[TextEffectFormat](texteffectformat-object-publisher.md)** object that represents the text formatting properties of a WordArt object.


## Syntax

 _expression_. **TextEffect**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example adds a WordArt object to the active publication and formats and inserts additional into it.


```vb
Sub AddFormatNewWordArt() 
 With ActiveDocument.Pages(1).Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Snap ITC", FontSize:=30, FontBold:=msoTrue, _ 
 FontItalic:=msoFalse, Left:=150, Top:=130) 
 .Rotation = 90 
 With .TextEffect 
 .RotatedChars = msoTrue 
 .Text = "This is a " &; .Text 
 End With 
 .Width = 250 
 End With 
End Sub
```


