---
title: Shape.GroupItems Property (Publisher)
keywords: vbapb10.chm2228280
f1_keywords:
- vbapb10.chm2228280
ms.prod: publisher
api_name:
- Publisher.Shape.GroupItems
ms.assetid: 9194f43b-bd8a-76a9-aa8c-17544d052d47
ms.date: 06/08/2017
---


# Shape.GroupItems Property (Publisher)

Returns a  **[GroupShapes](groupshapes-object-publisher.md)** collection if the specified shape is a group.


## Syntax

 _expression_. **GroupItems**

 _expression_A variable that represents a  **Shape** object.


## Remarks

All smart objects will be treated as grouped shapes.


## Example

This example adds three triangles to a publication, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


```vb
Sub Grouper() 
 
 Dim docSheet As Document 
 
 Set docSheet = ActiveDocument 
 With docSheet.MasterPages.Item(1).Shapes 
 ' Add the 3 triangles 
 .AddShape(Type:=msoShapeIsoscelesTriangle, _ 
 Left:=10, Top:=10, Width:=100, Height:=100).Name = "shpOne" 
 .AddShape(Type:=msoShapeIsoscelesTriangle, _ 
 Left:=150, Top:=10, Width:=100, Height:=100).Name = "shpTwo" 
 .AddShape(Type:=msoShapeIsoscelesTriangle, _ 
 Left:=300, Top:=10, Width:=100, Height:=100).Name = "shpThree" 
 ' Group and fill the 3 triangles 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured msoTextureBlueTissuePaper 
 .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble 
 End With 
 End With 
 
End Sub
```


