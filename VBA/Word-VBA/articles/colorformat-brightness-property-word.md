---
title: ColorFormat.Brightness Property (Word)
keywords: vbawd10.chm163971273
f1_keywords:
- vbawd10.chm163971273
ms.prod: word
api_name:
- Word.ColorFormat.Brightness
ms.assetid: 3a184574-24dc-2ea2-24f2-ba0b0b06df2e
ms.date: 06/08/2017
---


# ColorFormat.Brightness Property (Word)

Returns a  **Single** that represents the brightness of a specified shape color. Read/write.


## Syntax

 _expression_ . **Brightness**

 _expression_ A variable that represents a **[ColorFormat](colorformat-object-word.md)** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **Brightness** property, 0 (zero) being neutral.


## Example

The following code example creates a new shape in the active document, sets the fill color, and brightens the color.


```vb
Sub NewTintedShape() 
 Dim shpHeart As Shape 
 
 Set shpHeart = ActiveDocument.Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=150, _ 
 Top:=150, Width:=250, Height:=250) 
 With shpHeart.Fill.ForeColor 
 .RGB = RGB(Red:=255, Green:=28, Blue:=0) 
 .Brightness = 0.4 
 End With 
End Sub
```


## See also


#### Concepts


[ColorFormat Object](colorformat-object-word.md)

