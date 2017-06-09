---
title: ColorFormat.Brightness Property (PowerPoint)
keywords: vbapp10.chm506007
f1_keywords:
- vbapp10.chm506007
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat.Brightness
ms.assetid: 5140244e-d70b-f764-c127-917506b4074d
ms.date: 06/08/2017
---


# ColorFormat.Brightness Property (PowerPoint)

Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write  **Single**.


## Syntax

 _expression_. **Brightness**

 _expression_ A variable that represents a **ColorFormat** object.


## Example

The following example sets the brightness for shape one on  `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1)

myDocument.Shapes(1).PictureFormat.Brightness = 0.3
```


## See also


#### Concepts


[ColorFormat Object](colorformat-object-powerpoint.md)

