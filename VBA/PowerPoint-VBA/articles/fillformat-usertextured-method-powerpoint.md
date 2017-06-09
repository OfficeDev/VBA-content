---
title: FillFormat.UserTextured Method (PowerPoint)
keywords: vbapp10.chm552010
f1_keywords:
- vbapp10.chm552010
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.UserTextured
ms.assetid: 351d00db-4ed3-6975-e9c6-4174e796395d
ms.date: 06/08/2017
---


# FillFormat.UserTextured Method (PowerPoint)

Fills the specified shape with small tiles of an image. 


## Syntax

 _expression_. **UserTextured**( **_TextureFile_** )

 _expression_ A variable that represents an **FillFormat** object.


## Remarks

If you want to fill the shape with one large image, use the  **UserPicture** method.


## Example

This example adds two rectangles to  `myDocument`. The rectangle on the left is filled with one large image of the picture in Tiles.bmp; the rectangle on the right is filled with many small tiles of the picture in Tiles.bmp


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill _
        .UserPicture "c:\windows\tiles.bmp"

    .AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill _
        .UserTextured "c:\windows\tiles.bmp"

End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

