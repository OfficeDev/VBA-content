---
title: Shape.BlackWhiteMode Property (PowerPoint)
keywords: vbapp10.chm547017
f1_keywords:
- vbapp10.chm547017
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.BlackWhiteMode
ms.assetid: bed5df5a-87b5-5e61-6d28-48a7776d0d83
ms.date: 06/08/2017
---


# Shape.BlackWhiteMode Property (PowerPoint)

Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write.


## Syntax

 _expression_. **BlackWhiteMode**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoBlackWhiteMode


## Remarks

The value of the  **BlackWhiteMode** property can be one of these **MsoBlackWhiteMode** constants.


||
|:-----|
|<strong>msoBlackWhiteAutomatic</strong>|
|
<strong>msoBlackWhiteBlack</strong>|
|
<strong>msoBlackWhiteBlackTextAndLine</strong>|
|
<strong>msoBlackWhiteDontShow</strong>|
|
<strong>msoBlackWhiteGrayOutline</strong>|
|
<strong>msoBlackWhiteGrayScale</strong>|
|
<strong>msoBlackWhiteHighContrast</strong>|
|
<strong>msoBlackWhiteInverseGrayScale</strong>|
|
<strong>msoBlackWhiteLightGrayScale</strong>|
|
<strong>msoBlackWhiteMixed</strong>|
|
<strong>msoBlackWhiteWhite</strong>|

## Example

This example sets shape one on  `myDocument` to appear in black-and-white mode. When you view the presentation in black-and-white mode, shape one will appear black, regardless of what color it is in color mode.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).BlackWhiteMode = msoBlackWhiteBlack
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

