---
title: ShadowFormat.Visible Property (PowerPoint)
keywords: vbapp10.chm554010
f1_keywords:
- vbapp10.chm554010
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.Visible
ms.assetid: 83508398-55b9-8ac4-1724-f97247006664
ms.date: 06/08/2017
---


# ShadowFormat.Visible Property (PowerPoint)

Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **ShadowFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Visible** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified object or object formatting is not visible.|
|**msoTrue**| The specified object or object formatting is visible.|

## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on the first slide in the active presentation. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it and makes it visible.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Shadow

    .Visible = msoTrue

    .OffsetX = 5

    .OffsetY = -3

End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-powerpoint.md)

