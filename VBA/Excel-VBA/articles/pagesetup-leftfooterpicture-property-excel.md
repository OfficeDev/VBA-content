---
title: PageSetup.LeftFooterPicture Property (Excel)
keywords: vbaxl10.chm473109
f1_keywords:
- vbaxl10.chm473109
ms.prod: excel
api_name:
- Excel.PageSetup.LeftFooterPicture
ms.assetid: 296aa5d6-0354-741a-f96a-fb88e4c2e9de
ms.date: 06/08/2017
---


# PageSetup.LeftFooterPicture Property (Excel)

Returns a  **[Graphic](graphic-object-excel.md)** object that represents the picture for the left section of the footer. Used to set attributes about the picture.


## Syntax

 _expression_ . **LeftFooterPicture**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

The  **LeftFooterPicture** property is read-only, but the properties on it are not all read-only.


 **Note**  It is required that "&;G" is a part of the  **LeftFooter** property string in order for the image to show up in the left footer.


## Example

The following example adds a picture titled: Sample.jpg from the C:\ drive to the left section of the footer. This example assumes that a file called Sample.jpg exists on the C:\ drive.


```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.LeftFooterPicture 
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
 
 ' Enable the image to show up in the left footer. 
 ActiveSheet.PageSetup.LeftFooter = "&;G" 
 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

