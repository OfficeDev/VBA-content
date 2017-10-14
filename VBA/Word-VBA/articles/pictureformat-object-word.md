---
title: PictureFormat Object (Word)
keywords: vbawd10.chm2507
f1_keywords:
- vbawd10.chm2507
ms.prod: word
api_name:
- Word.PictureFormat
ms.assetid: 79556e36-81bb-f8df-45ef-c040df709497
ms.date: 06/08/2017
---


# PictureFormat Object (Word)

Contains properties and methods that apply to pictures and OLE objects. The  **LinkFormat** object contains properties and methods that apply to linked OLE objects only. The **OLEFormat** object contains properties and methods that apply to OLE objects whether or not they're linked.


## Remarks

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on the active document and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
With ActiveDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayScale 
 .CropBottom = 18 
End With
```


## Methods



|**Name**|
|:-----|
|[IncrementBrightness](pictureformat-incrementbrightness-method-word.md)|
|[IncrementContrast](pictureformat-incrementcontrast-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](pictureformat-application-property-word.md)|
|[Brightness](pictureformat-brightness-property-word.md)|
|[ColorType](pictureformat-colortype-property-word.md)|
|[Contrast](pictureformat-contrast-property-word.md)|
|[Creator](pictureformat-creator-property-word.md)|
|[Crop](pictureformat-crop-property-word.md)|
|[CropBottom](pictureformat-cropbottom-property-word.md)|
|[CropLeft](pictureformat-cropleft-property-word.md)|
|[CropRight](pictureformat-cropright-property-word.md)|
|[CropTop](pictureformat-croptop-property-word.md)|
|[Parent](pictureformat-parent-property-word.md)|
|[TransparencyColor](pictureformat-transparencycolor-property-word.md)|
|[TransparentBackground](pictureformat-transparentbackground-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
