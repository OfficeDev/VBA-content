---
title: ShapeRange.Flip Method (Word)
keywords: vbawd10.chm162856975
f1_keywords:
- vbawd10.chm162856975
ms.prod: word
api_name:
- Word.ShapeRange.Flip
ms.assetid: 363c222b-f0fc-8d42-5b06-82ec607a00c7
ms.date: 06/08/2017
---


# ShapeRange.Flip Method (Word)

Flips a shape horizontally or vertically.


## Syntax

 _expression_ . **Flip**( **_FlipCmd_** )

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FlipCmd_|Required| **MsoFlipCmd**|The flip orientation.|

## Example

This example adds a triangle to the active document, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.


```vb
Sub FlipShape() 
 With ActiveDocument.Shapes.AddShape( _ 
 Type:=msoShapeRightTriangle, Left:=150, _ 
 Top:=150, Width:=50, Height:=50).Duplicate 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .Flip msoFlipVertical 
 End With 
End Sub
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

