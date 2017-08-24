---
title: ThreeDFormat.ExtrusionColor Property (Publisher)
keywords: vbapb10.chm3801345
f1_keywords:
- vbapb10.chm3801345
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.ExtrusionColor
ms.assetid: 209a47fd-a219-9533-1a4a-572dfa4312f2
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColor Property (Publisher)

Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing the color of the shape's extrusion.


## Syntax

 _expression_. **ExtrusionColor**

 _expression_A variable that represents an  **ThreeDFormat** object.


### Return Value

ColorFormat


## Example

This example adds an oval to the active publication, and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Dim shpNew As Shape 
 
' Set a reference to a new oval. 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=90, Top:=90, Width:=90, Height:=40) 
 
' Format the 3-D properties of the oval. 
With shpNew.ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With 

```


