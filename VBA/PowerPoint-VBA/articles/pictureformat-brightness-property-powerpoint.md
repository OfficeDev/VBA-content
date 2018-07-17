---
title: PictureFormat.Brightness Property (PowerPoint)
keywords: vbapp10.chm551004
f1_keywords:
- vbapp10.chm551004
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.Brightness
ms.assetid: 11c01089-a69a-4ad0-ec01-b8d47a9f63f3
ms.date: 06/08/2017
---


# PictureFormat.Brightness Property (PowerPoint)

Returns or sets the brightness of the specified picture or OLE object. Read/write.


## Syntax

 _expression_. **Brightness**

 _expression_ A variable that represents a **PictureFormat** object.


### Return Value

Single


## Remarks

The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). 


## Example

This example sets the brightness for shape one on  `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).PictureFormat.Brightness = 0.3
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

