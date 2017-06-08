---
title: PictureFormat.CropLeft Property (Word)
keywords: vbawd10.chm164298856
f1_keywords:
- vbawd10.chm164298856
ms.prod: word
api_name:
- Word.PictureFormat.CropLeft
ms.assetid: c20c723a-c09b-f821-4273-9a5fc0f37207
ms.date: 06/08/2017
---


# PictureFormat.CropLeft Property (Word)

Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write  **Single** .


## Syntax

 _expression_ . **CropLeft**

 _expression_ A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it is 200 points wide, and then set the  **CropLeft** property to 50, 100 points (not 50) will be cropped off the left side of your picture.


## Example

This example crops 20 points off the left side of shape three on the active document. For the example to work, shape three must be either a picture or an OLE object.


```vb
ActiveDocument.Shapes(3).PictureFormat.CropLeft = 20
```

This example crops the percentage specified by the user off the left side of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim dblPercent As Double 
Dim shapeCrop As Shape 
Dim sngHeight As Single 
Dim sngCrop As Single 
 
dblPercent = Val(InputBox("What percentage do you want " _ 
 &; "to crop off the left of this picture?")) 
 
Set shapeCrop = _ 
 Selection.ShapeRange(1) 
 
With shapeCrop.Duplicate 
 .ScaleHeight 1, True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngCrop = sngHeight * dblPercent / 100 
 
shapeCrop.PictureFormat.CropLeft = sngCrop
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

