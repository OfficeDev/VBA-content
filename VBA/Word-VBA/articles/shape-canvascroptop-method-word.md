---
title: Shape.CanvasCropTop Method (Word)
keywords: vbawd10.chm161480845
f1_keywords:
- vbawd10.chm161480845
ms.prod: word
api_name:
- Word.Shape.CanvasCropTop
ms.assetid: 5d7cd546-4d47-d45a-9642-b2cd1e955480
ms.date: 06/08/2017
---


# Shape.CanvasCropTop Method (Word)

Crops a percentage of the height of a drawing canvas from the top of the canvas.


## Syntax

 _expression_ . **CanvasCropBottom**( **_Increment_** )

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|The amount in percentage points of a canvas's height that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's height from the top. Entering 0.1 crops ninety percent of the canvas's height from the top.|

## Example

This example crops twenty-five percent of the drawing canvas's height from the top of the first canvas in the active document, assuming the first shape in the active document is a drawing canvas. If not, you will need to add a drawing canvas to the document using the AddCanvas method.


```vb
Sub CropCanvasTop() 
 Dim shpCanvas As Shape 
 
 Set shpCanvas = ActiveDocument.Shapes(1) 
 shpCanvas.CanvasCropTop Increment:=0.75 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

