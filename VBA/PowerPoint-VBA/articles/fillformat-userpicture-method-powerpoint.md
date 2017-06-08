---
title: FillFormat.UserPicture Method (PowerPoint)
keywords: vbapp10.chm552009
f1_keywords:
- vbapp10.chm552009
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.UserPicture
ms.assetid: 87f28942-a5d2-7e27-7eee-5181d112d6d2
ms.date: 06/08/2017
---


# FillFormat.UserPicture Method (PowerPoint)

Fills the specified shape with one large image. 


## Syntax

 _expression_. **UserPicture**( **_PictureFile_** )

 _expression_ A variable that represents an **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PictureFile_|Required|**String**|The name of the picture file.|

## Remarks

If you want to fill the shape with small tiles of an image, use the  **[UserTextured](fillformat-usertextured-method-powerpoint.md)** method.


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

