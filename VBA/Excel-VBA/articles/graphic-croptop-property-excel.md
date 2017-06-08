---
title: Graphic.CropTop Property (Excel)
keywords: vbaxl10.chm694079
f1_keywords:
- vbaxl10.chm694079
ms.prod: excel
api_name:
- Excel.Graphic.CropTop
ms.assetid: fd35796f-6a5b-b914-d265-ab6bfc740981
ms.date: 06/08/2017
---


# Graphic.CropTop Property (Excel)

Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write  **Single** .


## Syntax

 _expression_ . **CropTop**

 _expression_ An expression that returns a **Graphic** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it's 200 points high, and then set the  **CropTop** property to 50, 100 points (not 50) will be cropped off the top of your picture.


## Example

This example crops 20 points off the top of shape three on  `myDocument`. For the example to work, shape three must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).PictureFormat.CropTop = 20
```

This example allows you to specify the percentage you want to crop off the top of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
percentToCrop = InputBox( _ 
 "What percentage do you want to crop" &; _ 
 " off the top of this picture?") 
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


[Graphic Object](graphic-object-excel.md)

