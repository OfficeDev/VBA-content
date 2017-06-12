---
title: Shape.CanvasCropBottom Method (Word)
keywords: vbawd10.chm161480847
f1_keywords:
- vbawd10.chm161480847
ms.prod: word
api_name:
- Word.Shape.CanvasCropBottom
ms.assetid: 13e9d954-3f95-2cf1-e2d7-314b67e25e33
ms.date: 06/08/2017
---


# Shape.CanvasCropBottom Method (Word)

Crops a percentage of the height of a drawing canvas from the bottom of the canvas.


## Syntax

 _expression_ . **CanvasCropBottom**( **_Increment_** )

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|The amount in percentage points of a drawing canvas's height that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's height from the bottom. Entering 0.1 crops ninety percent of the canvas's height from the bottom.|

## Example

This example crops twenty-five percent of the drawing canvas's height from the bottom of the first canvas in the active document, assuming the first shape in the active document is a drawing canvas. If not, you will need to add a drawing canvas to the document using the AddCanvas method.


```vb
Sub CropCanvasBottom() 
 Dim shpCanvas As Shape 
 
 Set shpCanvas = ActiveDocument.Shapes(1) 
 shpCanvas.CanvasCropBottom Increment:=0.75 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

