---
title: Shape.PictureFormat Property (PowerPoint)
keywords: vbapp10.chm547032
f1_keywords:
- vbapp10.chm547032
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.PictureFormat
ms.assetid: 97d6b8d0-ddfb-c3b8-70fe-7569f5738f92
ms.date: 06/08/2017
---


# Shape.PictureFormat Property (PowerPoint)

Returns a  **[PictureFormat](pictureformat-object-powerpoint.md)** object that contains picture formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **PictureFormat**

 _expression_ A variable that represents a **Shape** object.


### Return Value

PictureFormat


## Remarks

This property applies to  **Shape** or **ShapeRange** objects that represent pictures or OLE objects.


## Example

This example sets the brightness and contrast for shape one on  `myDocument`. Shape one must be a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).PictureFormat

    .Brightness = 0.3

    .Contrast = .75

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

