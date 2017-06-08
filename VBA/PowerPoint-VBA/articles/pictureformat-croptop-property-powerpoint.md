---
title: PictureFormat.CropTop Property (PowerPoint)
keywords: vbapp10.chm551010
f1_keywords:
- vbapp10.chm551010
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.CropTop
ms.assetid: dc9ef14a-99e0-6d5d-3df8-d7818569f31a
ms.date: 06/08/2017
---


# PictureFormat.CropTop Property (PowerPoint)

Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write. 


## Syntax

 _expression_. **CropTop**

 _expression_ A variable that represents a **PictureFormat** object.


### Return Value

Single


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it is 200 points high, and then set the  **CropTop** property to 50, 100 points (not 50) will be cropped off the top of your picture.


## Example

This example crops 20 points off the top of shape three on  `myDocument`. For the example to work, shape three must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).PictureFormat.CropTop = 20
```

This example crops the percentage specified by the user off the top of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```
percentToCrop = InputBox("What percentage do you " &; _
    "want to crop off the top of this picture?")

Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1)

With shapeToCrop.Duplicate
    .ScaleHeight 1, True
    origHeight = .Height
    .Delete
End With

cropPoints = origHeight * percentToCrop / 100
shapeToCrop.PictureFormat.CropTop = cropPoints
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

