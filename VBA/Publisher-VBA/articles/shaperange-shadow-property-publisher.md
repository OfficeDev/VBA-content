---
title: ShapeRange.Shadow Property (Publisher)
keywords: vbapb10.chm2293832
f1_keywords:
- vbapb10.chm2293832
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Shadow
ms.assetid: d6ee257c-9a26-abfc-9e8e-ef89bf627690
ms.date: 06/08/2017
---


# ShapeRange.Shadow Property (Publisher)

Returns a  **[ShadowFormat](shadowformat-object-publisher.md)** object that represents the shadow formatting for the specified shape.


## Syntax

 _expression_. **Shadow**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example adds an arrow with shadow formatting and fill color to the first page in the active document.


```vb
Sub SetShapeShadow() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRightArrow, Left:=72, _ 
 Top:=72, Width:=64, Height:=43) 
 .Shadow.Type = msoShadow5 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=255) 
 End With 
End Sub
```


