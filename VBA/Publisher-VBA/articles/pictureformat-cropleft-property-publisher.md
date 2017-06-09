---
title: PictureFormat.CropLeft Property (Publisher)
keywords: vbapb10.chm3604740
f1_keywords:
- vbapb10.chm3604740
ms.prod: publisher
api_name:
- Publisher.PictureFormat.CropLeft
ms.assetid: f9fd2031-83f7-ea81-84eb-4f1ac6d65082
ms.date: 06/08/2017
---


# PictureFormat.CropLeft Property (Publisher)

Returns or sets a  **Variant** indicating the amount by which the left edge of a picture or OLE object is cropped. Read/write.


## Syntax

 _expression_. **CropLeft**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

Negative values crop the bottom edge away from the center of the frame and positive values crop toward the right edge of the frame.

The valid range of crop values depends on the frame's position and size. For an unrotated frame, the lowest negative value allowed is the distance between the left edge of frame and the left edge of the scratch area. The highest positive value allowed is the current frame width.

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it is 200 points wide, and then set the  **CropLeft** property to 50, 100 points (not 50) will be cropped off the left of your picture.

Use the  **[CropRight](pictureformat-cropright-property-publisher.md)**,  **[CropTop](pictureformat-croptop-property-publisher.md)**, and  **[CropBottom](pictureformat-cropbottom-property-publisher.md)** properties to crop other edges of a picture or OLE object.


## Example

This example crops 20 points off the left of the third shape in the active publication. For the example to work, the shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropLeft = 20
```

This example crops the percentage specified by the user off the left of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngWidth As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the left of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleWidth Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngWidth = .Width 
 .Delete 
End With 
 
sngPoints = sngWidth * sngPercent / 100 
 
shpCrop.PictureFormat.CropLeft = sngPoints 

```


