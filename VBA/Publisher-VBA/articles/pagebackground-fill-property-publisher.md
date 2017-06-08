---
title: PageBackground.Fill Property (Publisher)
keywords: vbapb10.chm8126467
f1_keywords:
- vbapb10.chm8126467
ms.prod: publisher
api_name:
- Publisher.PageBackground.Fill
ms.assetid: bb5226aa-0b47-0d0f-1310-5abb34999910
ms.date: 06/08/2017
---


# PageBackground.Fill Property (Publisher)

 Returns a **[FillFormat](fillformat-object-publisher.md)** object representing the fill for the specified shape or table cell.


## Syntax

 _expression_. **Fill**

 _expression_A variable that represents a  **PageBackground** object.


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


