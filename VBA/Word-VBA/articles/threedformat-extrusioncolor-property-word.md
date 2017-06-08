---
title: ThreeDFormat.ExtrusionColor Property (Word)
keywords: vbawd10.chm164626533
f1_keywords:
- vbawd10.chm164626533
ms.prod: word
api_name:
- Word.ThreeDFormat.ExtrusionColor
ms.assetid: 60c8bf56-1a6e-08e9-2100-058c7863e2fe
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColor Property (Word)

Returns a  **[ColorFormat](colorformat-object-word.md)** object that represents the color of the shape's extrusion. Read-only.


## Syntax

 _expression_ . **ExtrusionColor**

 _expression_ A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Example

This example adds an oval to the active document and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Dim docActive As Document 
Dim shapeNew As Shape 
 
Set docActive = ActiveDocument 
Set shapeNew = docActive.Shapes.AddShape(msoShapeOval, _ 
 90, 90, 90, 40) 
With shapeNew.ThreeD 
 .Visible = True 
 .Depth = 50 
 ' RGB value for purple 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

