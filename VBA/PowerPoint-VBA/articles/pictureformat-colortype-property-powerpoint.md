---
title: PictureFormat.ColorType Property (PowerPoint)
keywords: vbapp10.chm551005
f1_keywords:
- vbapp10.chm551005
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.ColorType
ms.assetid: 5760f2e0-2247-1414-d2df-83666ca0a3b2
ms.date: 06/08/2017
---


# PictureFormat.ColorType Property (PowerPoint)

Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write.


## Syntax

 _expression_. **ColorType**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

MsoPictureColorType


## Remarks

The value of the  **ColorType** property can be one of these **MsoPictureColorType** constants.


||
|:-----|
|**msoPictureAutomatic**|
|**msoPictureBlackAndWhite**|
|**msoPictureGrayscale**|
|**msoPictureMixed**|
|**msoPictureWatermark**|

## Example

This example sets the color transformation to grayscale for shape one on  `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).PictureFormat.ColorType = msoPictureGrayScale
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

