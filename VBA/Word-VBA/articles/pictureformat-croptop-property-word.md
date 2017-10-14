---
title: PictureFormat.CropTop Property (Word)
keywords: vbawd10.chm164298858
f1_keywords:
- vbawd10.chm164298858
ms.prod: word
api_name:
- Word.PictureFormat.CropTop
ms.assetid: 724fbcad-20e9-896f-c832-1105b4e4d4d0
ms.date: 06/08/2017
---


# PictureFormat.CropTop Property (Word)

Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write  **Single** .


## Syntax

 _expression_ . **CropTop**

 _expression_ A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it is 200 points high, and then set the  **CropTop** property to 50, 100 points (not 50) will be cropped off the top of your picture.


## Example

This example crops 20 points off the top of shape three on the active document. For the example to work, shape three must be either a picture or an OLE object.


```vb
ActiveDocument.Shapes(3).PictureFormat.CropTop = 20
```

This example crops the percentage specified by the user off the top of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim dblPercent As Double 
Dim shapeCrop As Shape 
Dim sngHeight As Single 
Dim sngCrop As Single 
 
dblPercent = Val(InputBox("What percentage do you want " _ 
 &; "to crop off the top of this picture?")) 
 
Set shapeCrop = _ 
 Selection.ShapeRange(1) 
 
With shapeCrop.Duplicate 
 .ScaleHeight 1, True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngCrop = sngHeight * dblPercent / 100 
 
shapeCrop.PictureFormat.CropTop = sngCrop
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

