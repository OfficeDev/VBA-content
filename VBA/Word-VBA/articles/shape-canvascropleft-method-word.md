---
title: Shape.CanvasCropLeft Method (Word)
keywords: vbawd10.chm161480844
f1_keywords:
- vbawd10.chm161480844
ms.prod: word
api_name:
- Word.Shape.CanvasCropLeft
ms.assetid: d25b5c0e-82e7-1a33-9889-7b3df7f22802
ms.date: 06/08/2017
---


# Shape.CanvasCropLeft Method (Word)

Crops a percentage of the width of a drawing canvas from the left side of the canvas.


## Syntax

 _expression_ . **CanvasCropBottom**( **_Increment_** )

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|The amount in percentage points of the drawing canvas's width that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's width from the left. Entering 0.1 crops ninety percent of the canvas's width from the left.|

## Example

This example crops twenty-five percent of the drawing canvas's width from the left side of the first canvas in the active document, assuming the first shape in the active document is a drawing canvas. If not, you will need to add a drawing canvas to the document using the AddCanvas method.


```vb
Sub CropCanvasLeft() 
 Dim shpCanvas As Shape 
 
 Set shpCanvas = ActiveDocument.Shapes(1) 
 shpCanvas.CanvasCropLeft Increment:=0.75 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

