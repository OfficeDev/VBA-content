---
title: DropCap.FontColor Property (Publisher)
keywords: vbapb10.chm5505028
f1_keywords:
- vbapb10.chm5505028
ms.prod: publisher
api_name:
- Publisher.DropCap.FontColor
ms.assetid: 0c740ec7-05ac-b1fc-875c-cfd5a934c403
ms.date: 06/08/2017
---


# DropCap.FontColor Property (Publisher)

Returns or sets a  **[ColorFormat](colorformat-object-publisher.md)** object that represents the color applied to a specified dropped capital letter.


## Syntax

 _expression_. **FontColor**

 _expression_A variable that represents a  **DropCap** object.


### Return Value

ColorFormat


## Example

This example applies an  **[RGB](colorformat-rgb-property-publisher.md)** color to the dropped capital letter in the specified text frame. This example assumes that the specified text frame is formatted with a dropped capital letter.


```vb
Sub BoldDropCap() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.DropCap 
 .FontBold = msoTrue 
 .FontColor.RGB = RGB(Red:=150, Green:=50, Blue:=180) 
 .FontItalic = msoTrue 
 .FontName = "Script MT Bold" 
 End With 
End Sub
```


