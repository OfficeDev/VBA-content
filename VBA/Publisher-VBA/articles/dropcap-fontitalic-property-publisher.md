---
title: DropCap.FontItalic Property (Publisher)
keywords: vbapb10.chm5505030
f1_keywords:
- vbapb10.chm5505030
ms.prod: publisher
api_name:
- Publisher.DropCap.FontItalic
ms.assetid: 57996a71-94db-67b0-ee64-bd79144d01d1
ms.date: 06/08/2017
---


# DropCap.FontItalic Property (Publisher)

Sets or returns an  **MsoTriState**constant that represents whether the font for a dropped capital letter or WordArt text effect is italic. Read/write.


## Syntax

 _expression_. **FontItalic**

 _expression_A variable that represents a  **DropCap** object.


## Remarks

The  **FontItalic** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example makes the dropped capital letter in the specified text frame italic. This example assumes that the specified text frame is formatted with a dropped capital letter.


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


