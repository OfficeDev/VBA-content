---
title: PictureFormat.CropRight Property (Word)
keywords: vbawd10.chm164298857
f1_keywords:
- vbawd10.chm164298857
ms.prod: word
api_name:
- Word.PictureFormat.CropRight
ms.assetid: 89f73474-9b52-b758-e579-adbc803a5a62
ms.date: 06/08/2017
---


# PictureFormat.CropRight Property (Word)

Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write  **Single** .


## Syntax

 _expression_ . **CropRight**

 _expression_ A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it is 200 points wide, and then set the  **CropRight** property to 50, 100 points (not 50) will be cropped off the right side of your picture.


## Example

This example crops 20 points off the right side of shape three on the active document. For this example to work, shape three must be either a picture or an OLE object.


```vb
ActiveDocument.Shapes(3).PictureFormat.CropRight = 20
```

This example crops the percentage specified by the user off the right side of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim dblPercent As Double 
Dim shapeCrop As Shape 
Dim sngHeight As Single 
Dim sngCrop As Single 
 
dblPercent = Val(InputBox("What percentage do you want " _ 
 &; "to crop off the right of this picture?")) 
 
Set shapeCrop = _ 
 Selection.ShapeRange(1) 
 
With shapeCrop.Duplicate 
 .ScaleHeight 1, True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngCrop = sngHeight * dblPercent / 100 
 
shapeCrop.PictureFormat.CropRight = sngCrop
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

