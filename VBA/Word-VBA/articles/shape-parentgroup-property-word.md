---
title: Shape.ParentGroup Property (Word)
keywords: vbawd10.chm161480841
f1_keywords:
- vbawd10.chm161480841
ms.prod: word
api_name:
- Word.Shape.ParentGroup
ms.assetid: c6305148-86d4-9f86-45e9-5007d7f5b324
ms.date: 06/08/2017
---


# Shape.ParentGroup Property (Word)

Returns a  **Shape** object that represents the common parent shape of a child shape or a range of child shapes.


## Syntax

 _expression_ . **ParentGroup**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example creates two shapes in the active document and groups those shapes. Then using one shape in the group, it accesses the parent group and fills all shapes in the parent group with the same fill color. This example assumes that the active document does not currently contain any shapes. If it does, an error may occur.


```vb
Sub ParentGroupShape() 
 Dim pgShape As Shape 
 
 'Add two shapes to active document and group 
 With ActiveDocument.Shapes 
 .AddShape Type:=msoShapeOval, Left:=72, _ 
 Top:=72, Width:=100, Height:=100 
 .AddShape Type:=msoShapeHeart, Left:=110, _ 
 Top:=120, Width:=100, Height:=100 
 .Range(Array(1, 2)).Group 
 End With 
 
 Set pgShape = ActiveDocument.Shapes(1) _ 
 .GroupItems(1).ParentGroup 
 pgShape.Fill.ForeColor.RGB = RGB(Red:=100, Green:=0, Blue:=255) 
 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

