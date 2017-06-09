---
title: PageSetup.CenterFooterPicture Property (Excel)
keywords: vbaxl10.chm473107
f1_keywords:
- vbaxl10.chm473107
ms.prod: excel
api_name:
- Excel.PageSetup.CenterFooterPicture
ms.assetid: 6df72e33-29d2-a638-7e42-2749a61ff9a3
ms.date: 06/08/2017
---


# PageSetup.CenterFooterPicture Property (Excel)

Returns a  **[Graphic](graphic-object-excel.md)** object that represents the picture for the center section of the footer. Used to set attributes about the picture.


## Syntax

 _expression_ . **CenterFooterPicture**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

The  **CenterFooterPicture** property is read-only, but the properties on it are not all read-only.

It is required that "&;G" is a part of the  **CenterFooter** property string in order for the image to show up in the center footer.


## Example

The following example adds a picture titled: Sample.jpg from the C:\ drive to the center section of the footer. This example assumes that a file called Sample.jpg exists on the C:\ drive.


```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.CentertFooterPicture 
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
 
 ' Enable the image to show up in the center footer. 
 ActiveSheet.PageSetup.CenterFooter = "&;G" 
 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

