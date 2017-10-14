---
title: PictureFormat.CropBottom Property (Word)
keywords: vbawd10.chm164298855
f1_keywords:
- vbawd10.chm164298855
ms.prod: word
api_name:
- Word.PictureFormat.CropBottom
ms.assetid: f7cf6d4a-cc95-f595-9382-1daf4e0cf8de
ms.date: 06/08/2017
---


# PictureFormat.CropBottom Property (Word)

Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write  **Single** .


## Syntax

 _expression_ . **CropBottom**

 _expression_ A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it is 200 points high, and then set the  **CropBottom** property to 50, 100 points (not 50) will be cropped off the bottom of your picture.


## Example

This example crops 20 points off the bottom of shape three on the active document. For the example to work, shape three must be either a picture or an OLE object.


```vb
ActiveDocument.Shapes(3).PictureFormat.CropBottom = 20
```

This example crops the percentage specified by the user off the bottom of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim dblPercent As Double 
Dim shapeCrop As Shape 
Dim sngHeight As Single 
Dim sngCrop As Single 
 
dblPercent = Val(InputBox("What percentage do you want " _ 
 &; "to crop off the bottom of this picture?")) 
 
Set shapeCrop = _ 
 Selection.ShapeRange(1) 
 
With shapeCrop.Duplicate 
 .ScaleHeight 1, True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngCrop = sngHeight * dblPercent / 100 
 
shapeCrop.PictureFormat.CropBottom = sngCrop
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

