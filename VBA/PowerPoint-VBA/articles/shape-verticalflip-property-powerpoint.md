---
title: Shape.VerticalFlip Property (PowerPoint)
keywords: vbapp10.chm547039
f1_keywords:
- vbapp10.chm547039
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.VerticalFlip
ms.assetid: 56bf36e4-49df-5ae5-855c-3275d634dee4
ms.date: 06/08/2017
---


# Shape.VerticalFlip Property (PowerPoint)

Determines whether the specified shape is flipped around the vertical axis. Read-only.


## Syntax

 _expression_. **VerticalFlip**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoTriState


## Remarks

The value of the  **VerticalFlip** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape is not flipped around the vertical axis.|
|**msoTrue**| The specified shape is flipped around the vertical axis.|

## Example

This example restores each shape on  `myDocument` to its original state if it is been flipped horizontally or vertically.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.HorizontalFlip Then s.Flip msoFlipHorizontal

    If s.VerticalFlip Then s.Flip msoFlipVertical

Next
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

