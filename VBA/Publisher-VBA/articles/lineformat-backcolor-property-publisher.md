---
title: LineFormat.BackColor Property (Publisher)
keywords: vbapb10.chm3408128
f1_keywords:
- vbapb10.chm3408128
ms.prod: publisher
api_name:
- Publisher.LineFormat.BackColor
ms.assetid: 45e18a2e-4354-65d7-9a80-53869c4914f0
ms.date: 06/08/2017
---


# LineFormat.BackColor Property (Publisher)

Returns or sets a  **[ColorFormat](colorformat-object-publisher.md)** object representing the background color for the specified fill or patterned line. .


## Syntax

 _expression_. **BackColor**

 _expression_A variable that represents a  **LineFormat** object.


## Remarks

Use the  **[ForeColor](fillformat-forecolor-property-publisher.md)** property to set the foreground color for a fill or line.


## Example

This example adds a rectangle to the active publication and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient _ 
 Style:=msoGradientHorizontal, Variant:=1 
End With
```

This example adds a patterned line to the active publication.




```vb
With ActiveDocument.Pages(1).Shapes.AddLine _ 
 (BeginX:=10, BeginY:=100, EndX:=250, EndY:=0).Line 
 .Weight = 6 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(128, 0, 0) 
 .Pattern = msoPatternDarkDownwardDiagonal 
End With 

```


