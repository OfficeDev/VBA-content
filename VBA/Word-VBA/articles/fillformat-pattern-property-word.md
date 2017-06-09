---
title: FillFormat.Pattern Property (Word)
keywords: vbawd10.chm164102250
f1_keywords:
- vbawd10.chm164102250
ms.prod: word
api_name:
- Word.FillFormat.Pattern
ms.assetid: a0a296b4-20d2-20a6-9892-e22d1b7f4cff
ms.date: 06/08/2017
---


# FillFormat.Pattern Property (Word)

Returns or sets a  **MsoPatternType** constant that represents the pattern applied to the specified fill or line. Read-only.


## Syntax

 _expression_ . **Pattern**

 _expression_ An expression that returns a **[FillFormat](fillformat-object-word.md)** object.


## Example

This example adds a rectangle to myDocument and sets its fill pattern to match that of the shape named "rect1." The new rectangle has the same pattern as "rect1" but not necessarily the same colors. The colors used in the pattern are set with the  **BackColor** and **ForeColor** properties.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes 
 pattern1 = .Item("rect1").Fill.Pattern 
 With .AddShape(msoShapeRectangle, 100, 100, 120, 80).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 0, 255) 
 .Patterned pattern1 
 End With 
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

