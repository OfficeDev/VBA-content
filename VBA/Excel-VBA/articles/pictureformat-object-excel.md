---
title: PictureFormat Object (Excel)
keywords: vbaxl10.chm113000
f1_keywords:
- vbaxl10.chm113000
ms.prod: excel
api_name:
- Excel.PictureFormat
ms.assetid: 7e8ec723-b6e0-fdc9-ff4e-22cbb31be4df
ms.date: 06/08/2017
---


# PictureFormat Object (Excel)

Contains properties and methods that apply to pictures and OLE objects.


## Remarks

 The **[LinkFormat](linkformat-object-excel.md)** object contains properties and methods that apply to linked OLE objects only. The **[OLEFormat](oleformat-object-excel.md)** object contains properties and methods that apply to OLE objects whether or not they're linked.


## Example

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on _myDocument_ and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayScale 
 .CropBottom = 18
```


## Methods



|**Name**|
|:-----|
|[IncrementBrightness](pictureformat-incrementbrightness-method-excel.md)|
|[IncrementContrast](pictureformat-incrementcontrast-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](pictureformat-application-property-excel.md)|
|[Brightness](pictureformat-brightness-property-excel.md)|
|[ColorType](pictureformat-colortype-property-excel.md)|
|[Contrast](pictureformat-contrast-property-excel.md)|
|[Creator](pictureformat-creator-property-excel.md)|
|[Crop](pictureformat-crop-property-excel.md)|
|[CropBottom](pictureformat-cropbottom-property-excel.md)|
|[CropLeft](pictureformat-cropleft-property-excel.md)|
|[CropRight](pictureformat-cropright-property-excel.md)|
|[CropTop](pictureformat-croptop-property-excel.md)|
|[Parent](pictureformat-parent-property-excel.md)|
|[TransparencyColor](pictureformat-transparencycolor-property-excel.md)|
|[TransparentBackground](pictureformat-transparentbackground-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
