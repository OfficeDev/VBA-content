---
title: PageSetup.CenterHeaderPicture Property (Excel)
keywords: vbaxl10.chm473106
f1_keywords:
- vbaxl10.chm473106
ms.prod: excel
api_name:
- Excel.PageSetup.CenterHeaderPicture
ms.assetid: c4c6e0b5-96e3-eaea-2dfe-807f286029ec
ms.date: 06/08/2017
---


# PageSetup.CenterHeaderPicture Property (Excel)

Returns a  **[Graphic](graphic-object-excel.md)** object that represents the picture for the center section of the header. Used to set attributes about the picture.


## Syntax

 _expression_ . **CenterHeaderPicture**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

The  **CenterHeaderPicture** property is read-only, but the properties on it are not all read-only.

It is required that "&;G" is a part of the  **CenterHeader** property string in order for the image to show up in the center header.


## Example

The following example adds a picture titled: Sample.jpg from the C:\ drive to the center section of the header. This example assumes that a file called Sample.jpg exists on the C:\ drive.


```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.CentertHeaderPicture 
 .FileName = "C:\Sample.jpg" 
 .Height = 275.25 
 .Width = 463.5 
 .Brightness = 0.36 
 .ColorType = msoPictureGrayscale 
 .Contrast = 0.39 
 .CropBottom = -14.4 
 .CropLeft = -28.8 
 .CropRight = -14.4 
 .CropTop = 21.6 
 End With 
 
 ' Enable the image to show up in the center header. 
 ActiveSheet.PageSetup.CenterHeader = "&;G" 
 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

