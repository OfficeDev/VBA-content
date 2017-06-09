---
title: PictureFormat.CropRight Property (PowerPoint)
keywords: vbapp10.chm551009
f1_keywords:
- vbapp10.chm551009
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.CropRight
ms.assetid: 217691ed-5533-707c-338d-4375dbdd3eaa
ms.date: 06/08/2017
---


# PictureFormat.CropRight Property (PowerPoint)

Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write. 


## Syntax

 _expression_. **CropRight**

 _expression_ A variable that represents a **PictureFormat** object.


### Return Value

Single


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it is 200 points wide, and then set the  **CropRight** property to 50, 100 points (not 50) will be cropped off the right side of your picture.


## Example

This example crops 20 points off the right side of shape three on  `myDocument`. For this example to work, shape three must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).PictureFormat.CropRight = 20
```

This example crops the percentage specified by the user off the right side of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```
percentToCrop = InputBox("What percentage do you " &; _
    "want to crop off the right of this picture?")

Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1)

With shapeToCrop.Duplicate
    .ScaleWidth 1, True
    origWidth = .Width
    .Delete
End With

cropPoints = origWidth * percentToCrop / 100

shapeToCrop.PictureFormat.CropRight = cropPoints
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

