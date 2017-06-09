---
title: ThreeDFormat.Depth Property (Word)
keywords: vbawd10.chm164626532
f1_keywords:
- vbawd10.chm164626532
ms.prod: word
api_name:
- Word.ThreeDFormat.Depth
ms.assetid: 45fbea95-7685-d244-19b8-ef4c4560a26f
ms.date: 06/08/2017
---


# ThreeDFormat.Depth Property (Word)

Returns or sets the depth of the shape's extrusion. Read/write  **Single** .


## Syntax

 _expression_ . **Depth**

 _expression_ A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Remarks

The  **Depth** property can be a value from - 600 through 9600 (positive values produce an extrusion whose front face is the original shape; negative values produce an extrusion whose back face is the original shape).


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

