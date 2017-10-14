---
title: PictureFormat Object (PowerPoint)
keywords: vbapp10.chm551000
f1_keywords:
- vbapp10.chm551000
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat
ms.assetid: 946794b4-0401-ec7c-cea3-779ebfce0d69
ms.date: 06/08/2017
---


# PictureFormat Object (PowerPoint)

Contains properties and methods that apply to pictures and OLE objects. 


## Example

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on `myDocument` and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).PictureFormat

    .Brightness = 0.3

    .Contrast = 0.7

    .ColorType = msoPictureGrayScale

    .CropBottom = 18

End With
```


## Methods



|**Name**|
|:-----|
|[IncrementBrightness](http://msdn.microsoft.com/library/4237d547-2c8b-9ed2-f131-6a4fb52ee0a2%28Office.15%29.aspx)|
|[IncrementContrast](http://msdn.microsoft.com/library/ad5c45b2-0193-eda9-a511-4dd9050daee7%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/1fe92d27-cc82-60be-b9b2-d1dbded71d5a%28Office.15%29.aspx)|
|[Brightness](http://msdn.microsoft.com/library/11c01089-a69a-4ad0-ec01-b8d47a9f63f3%28Office.15%29.aspx)|
|[ColorType](http://msdn.microsoft.com/library/5760f2e0-2247-1414-d2df-83666ca0a3b2%28Office.15%29.aspx)|
|[Contrast](http://msdn.microsoft.com/library/19e2a7d2-59c3-e3d7-3770-0cbecdba2550%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/d2784238-bf55-0e70-a89b-0a3c9b21fd31%28Office.15%29.aspx)|
|[Crop](http://msdn.microsoft.com/library/8e39ec71-ae5e-99a0-c090-a55d15c6e9f7%28Office.15%29.aspx)|
|[CropBottom](http://msdn.microsoft.com/library/6d2252ab-33ed-802b-e0c5-3e12be23bec4%28Office.15%29.aspx)|
|[CropLeft](http://msdn.microsoft.com/library/401a863f-9162-a8d8-825c-f615e6d25907%28Office.15%29.aspx)|
|[CropRight](http://msdn.microsoft.com/library/217691ed-5533-707c-338d-4375dbdd3eaa%28Office.15%29.aspx)|
|[CropTop](http://msdn.microsoft.com/library/dc9ef14a-99e0-6d5d-3df8-d7818569f31a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/bb25c345-3b60-0484-1c21-4f2af88cc20f%28Office.15%29.aspx)|
|[TransparencyColor](http://msdn.microsoft.com/library/122e69f6-a403-92d1-8ef7-087c9396ed6a%28Office.15%29.aspx)|
|[TransparentBackground](http://msdn.microsoft.com/library/b4a15c64-0568-dcd7-99a2-00295bfe679c%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
