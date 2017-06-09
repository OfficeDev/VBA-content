---
title: ShapeRange.CanvasCropRight Method (Word)
keywords: vbawd10.chm162857102
f1_keywords:
- vbawd10.chm162857102
ms.prod: word
api_name:
- Word.ShapeRange.CanvasCropRight
ms.assetid: c4000f8e-3dac-ae4d-3ed3-91c2ac471fa5
ms.date: 06/08/2017
---


# ShapeRange.CanvasCropRight Method (Word)

Crops a percentage of the width of a drawing canvas from the right side of the canvas.


## Syntax

 _expression_ . **CanvasCropBottom**( **_Increment_** )

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|The amount in percentage points of the canvas's width that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's width from the right. Entering 0.1 crops ninety percent of the canvas's width from the right.|

## Example

This example crops twenty-five percent of the drawing canvas's width from the right side of the first canvas in the active document, assuming the first shape in the active document is a drawing canvas. If not, you will need to add a drawing canvas to the document using the AddCanvas method.


```vb
Sub CropCanvasRight() 
 Dim shpCanvas As Shape 
 
 Set shpCanvas = ActiveDocument.Shapes(1) 
 shpCanvas.CanvasCropRight Increment:=0.75 
End Sub
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

