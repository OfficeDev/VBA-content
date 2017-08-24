---
title: Cell.Fill Property (Publisher)
keywords: vbapb10.chm5111817
f1_keywords:
- vbapb10.chm5111817
ms.prod: publisher
api_name:
- Publisher.Cell.Fill
ms.assetid: 3ff3fda8-aca7-534e-6a56-99d6a3d9e9e2
ms.date: 06/08/2017
---


# Cell.Fill Property (Publisher)

 Returns a **[FillFormat](fillformat-object-publisher.md)** object representing the fill for the specified shape or table cell.


## Syntax

 _expression_. **Fill**

 _expression_A variable that represents a  **Cell** object.


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


