---
title: FillFormat.Pattern Property (PowerPoint)
keywords: vbapp10.chm552017
f1_keywords:
- vbapp10.chm552017
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.Pattern
ms.assetid: 843504d6-d9a5-f732-89eb-d2d3d1ea4477
ms.date: 06/08/2017
---


# FillFormat.Pattern Property (PowerPoint)

Sets or returns a value that represents the pattern applied to the specified fill. Read-only.


## Syntax

 _expression_. **Pattern**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

[MsoPatternType](http://msdn.microsoft.com/library/b95a7e43-329f-b93b-3664-04d8f570c747%28Office.15%29.aspx)


## Remarks

Use the  **[BackColor](fillformat-backcolor-property-powerpoint.md)** and **[ForeColor](fillformat-forecolor-property-powerpoint.md)** properties to set the colors used in the pattern.


## Example

This example adds a rectangle to  `myDocument` and sets its fill pattern to match that of the shape named "rect1." The new rectangle has the same pattern as rect1, but not necessarily the same colors. The colors used in the pattern are set with the **BackColor** and **ForeColor** properties.


```vb
Set myDocument = ActivePresentation.Slides(1) 
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


[FillFormat Object](fillformat-object-powerpoint.md)

