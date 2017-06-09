---
title: Shape.HorizontalFlip Property (PowerPoint)
keywords: vbapp10.chm547025
f1_keywords:
- vbapp10.chm547025
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.HorizontalFlip
ms.assetid: bf061a08-978c-dfb3-8a8f-4ecd62d95c53
ms.date: 06/08/2017
---


# Shape.HorizontalFlip Property (PowerPoint)

Returns whether the specified shape is flipped around the horizontal axis. Read-only.


## Syntax

 _expression_. **HorizontalFlip**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HorizontalFlip** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The specified shape is not flipped around the horizontal axis.|
|**msoTrue**| The specified shape is flipped around the horizontal axis.|

## Example

This example restores each shape on  `myDocument` to its original state, if it is been flipped horizontally or vertically.


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

