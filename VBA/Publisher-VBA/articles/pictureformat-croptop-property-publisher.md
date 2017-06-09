---
title: PictureFormat.CropTop Property (Publisher)
keywords: vbapb10.chm3604742
f1_keywords:
- vbapb10.chm3604742
ms.prod: publisher
api_name:
- Publisher.PictureFormat.CropTop
ms.assetid: b235898d-addf-6a4c-5693-229431545e6c
ms.date: 06/08/2017
---


# PictureFormat.CropTop Property (Publisher)

Returns or sets a  **Variant** indicating the amount by which the top edge of a picture or OLE object is cropped. Read/write.


## Syntax

 _expression_. **CropTop**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

Negative values crop the top edge away from the center of the frame and positive values crop toward the bottom edge of the frame.

The valid range of crop values depends on the frame's position and size. For an unrotated frame, the lowest negative value allowed is the distance between the top edge of frame and the top edge of the scratch area. The highest positive value allowed is the current frame height.

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it is 200 points high, and then set the  **CropTop** property to 50, 100 points (not 50) will be cropped off the top of your picture.

Use the  **[CropLeft](pictureformat-cropleft-property-publisher.md)**,  **[CropRight](pictureformat-cropright-property-publisher.md)**, and  **[CropBottom](pictureformat-cropbottom-property-publisher.md)** properties to crop other edges of a picture or OLE object.


## Example

This example crops 20 points off the top of the third shape in the active publication. For the example to work, the shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropTop = 20
```

This example crops the percentage specified by the user off the top of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngHeight As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the top of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleHeight Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngPoints = sngHeight * sngPercent / 100 
 
shpCrop.PictureFormat.CropTop = sngPoints 

```


