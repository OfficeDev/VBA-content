---
title: Shape.Fill Property (Publisher)
keywords: vbapb10.chm2228279
f1_keywords:
- vbapb10.chm2228279
ms.prod: publisher
api_name:
- Publisher.Shape.Fill
ms.assetid: ff1b8d02-150e-e023-2f0a-b1608cc99644
ms.date: 06/08/2017
---


# Shape.Fill Property (Publisher)

 Returns a **[FillFormat](fillformat-object-publisher.md)** object representing the fill for the specified shape or table cell.


## Syntax

 _expression_. **Fill**

 _expression_A variable that represents a  **Shape** object.


## Example

This example creates a new  **AutoShape** object and fills the shape with green.


```vb
Sub NewShapeItem() 
 
 Dim shpHeart As Shape 
 
 Set shpHeart = ThisDocument.MasterPages.Item(1).Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=40, Top:=80, _ 
 Width:=50, Height:=50) 
 shpHeart.Fill.ForeColor.RGB = RGB(Red:=0, Green:=255, Blue:=0) 
 
End Sub
```


