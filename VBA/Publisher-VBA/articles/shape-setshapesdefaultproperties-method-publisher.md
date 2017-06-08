---
title: Shape.SetShapesDefaultProperties Method (Publisher)
keywords: vbapb10.chm2228264
f1_keywords:
- vbapb10.chm2228264
ms.prod: publisher
api_name:
- Publisher.Shape.SetShapesDefaultProperties
ms.assetid: 3f7d7143-3a08-6ff4-c28e-86598212a876
ms.date: 06/08/2017
---


# Shape.SetShapesDefaultProperties Method (Publisher)

Applies the formatting for the specified shape or shape range to the default shape. Shapes created after this method has been used will have this formatting applied to them by default.


## Syntax

 _expression_. **SetShapesDefaultProperties**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **SetShapesDefaultProperties** method stores two different sets of default properties, one for a **Shape** object's ** [AutoShapeType Property](shape-autoshapetype-property-publisher.md)**, and another for a  **TextFrame** object. In other words, if this method is called on an AutoShape, the default formatting of that object will apply only to new AutoShapes, and will not apply to new text boxes. If this method is called on a text box, the default formatting of that object will apply only to new text boxes, and will not apply to new AutoShapes.


## Example

This example adds a rectangle to the active publication, formats the rectangle's fill, applies the rectangle's formatting to the default shape, and then adds another smaller rectangle to the document. The second rectangle has the same fill as the first one.


```vb
With ActiveDocument.Pages(1).Shapes 
 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=5, Top:=5, Width:=80, Height:=60) 
 With .Fill 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(0, 204, 255) 
 .Patterned Pattern:=msoPatternHorizontalBrick 
 End With 
 .SetShapesDefaultProperties 
 End With 
 
 .AddShape Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=40, Height:=30 
 
End With 

```


