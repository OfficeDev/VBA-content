---
title: PictureFormat.CropRight Property (Excel)
keywords: vbaxl10.chm113007
f1_keywords:
- vbaxl10.chm113007
ms.prod: excel
api_name:
- Excel.PictureFormat.CropRight
ms.assetid: 9cf71268-5d63-4f66-6245-968786db14a8
ms.date: 06/08/2017
---


# PictureFormat.CropRight Property (Excel)

Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write  **Single** .


## Syntax

 _expression_ . **CropRight**

 _expression_ An expression that returns a **PictureFormat** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it's 200 points wide, and then set the  **CropRight** property to 50, 100 points (not 50) will be cropped off the right side of your picture.


## Example

This example crops 20 points off the right side of shape three on  `myDocument`. For this example to work, shape three must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).PictureFormat.CropRight = 20
```

Using this example, you can specify the percentage you want to cropp off the right side of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
percentToCrop = InputBox( _ 
 "What percentage do you want to crop" &; _ 
 " off the right of this picture?") 
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


[PictureFormat Object](pictureformat-object-excel.md)

