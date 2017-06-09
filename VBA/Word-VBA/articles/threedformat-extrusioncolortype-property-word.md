---
title: ThreeDFormat.ExtrusionColorType Property (Word)
keywords: vbawd10.chm164626534
f1_keywords:
- vbawd10.chm164626534
ms.prod: word
api_name:
- Word.ThreeDFormat.ExtrusionColorType
ms.assetid: cddfbdac-601b-1786-fe41-5d155114d539
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColorType Property (Word)

Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write  **MsoExtrusionColorType** .


## Syntax

 _expression_ . **ExtrusionColorType**

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Example

If the first shape on the active document has an automatic extrusion color, this example gives the extrusion a custom yellow color.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes(1).ThreeD 
 If .ExtrusionColorType = msoExtrusionColorAutomatic Then 
 .ExtrusionColor.RGB = RGB(240, 235, 16) 
 End If 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

