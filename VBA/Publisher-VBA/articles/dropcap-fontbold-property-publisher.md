---
title: DropCap.FontBold Property (Publisher)
keywords: vbapb10.chm5505029
f1_keywords:
- vbapb10.chm5505029
ms.prod: publisher
api_name:
- Publisher.DropCap.FontBold
ms.assetid: 7e1b9b51-258d-080c-e5ae-cdc9d6a2ba64
ms.date: 06/08/2017
---


# DropCap.FontBold Property (Publisher)

Sets or returns an  **MsoTriState**constant that represents whether the font for a dropped capital letter or WordArt text effect is bold. Read/write.


## Syntax

 _expression_. **FontBold**

 _expression_A variable that represents a  **DropCap** object.


## Remarks

The  **FontBold** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example applies bold formatting to the dropped capital letter in the specified text frame. This example assumes that the specified text frame is formatted with a dropped capital letter.


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


