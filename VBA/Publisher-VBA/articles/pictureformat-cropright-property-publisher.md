---
title: PictureFormat.CropRight Property (Publisher)
keywords: vbapb10.chm3604741
f1_keywords:
- vbapb10.chm3604741
ms.prod: publisher
api_name:
- Publisher.PictureFormat.CropRight
ms.assetid: b1c20de2-e2cf-708f-ddae-194c8b1b01c1
ms.date: 06/08/2017
---


# PictureFormat.CropRight Property (Publisher)

Returns or sets a  **Variant** indicating the amount by which the right edge of a picture or OLE object is cropped. Read/write.


## Syntax

 _expression_. **CropRight**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

Negative values crop the bottom edge away from the center of the frame and positive values crop toward the left edge of the frame.

The valid range of crop values depends on the frame's position and size. For an unrotated frame, the lowest negative value allowed is the distance between the right edge of frame and the right edge of the scratch area. The highest positive value allowed is the current frame width.

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it is 200 points wide, and then set the  **CropRight** property to 50, 100 points (not 50) will be cropped off the right of your picture.

Use the  **[CropLeft](pictureformat-cropleft-property-publisher.md)**,  **[CropTop](pictureformat-croptop-property-publisher.md)**, and  **[CropBottom](pictureformat-cropbottom-property-publisher.md)** properties to crop other edges of a picture or OLE object.


## Example

This example crops 20 points off the right of the third shape in the active publication. For the example to work, the shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropRight = 20
```

This example crops the percentage specified by the user off the right of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngWidth As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the right of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleWidth Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngWidth = .Width 
 .Delete 
End With 
 
sngPoints = sngWidth * sngPercent / 100 
 
shpCrop.PictureFormat.CropRight = sngPoints 

```


