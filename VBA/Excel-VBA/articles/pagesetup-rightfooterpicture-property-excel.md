---
title: PageSetup.RightFooterPicture Property (Excel)
keywords: vbaxl10.chm473111
f1_keywords:
- vbaxl10.chm473111
ms.prod: excel
api_name:
- Excel.PageSetup.RightFooterPicture
ms.assetid: f33bbfb1-91d0-6724-0944-2b63c6720d86
ms.date: 06/08/2017
---


# PageSetup.RightFooterPicture Property (Excel)

Returns a  **[Graphic](graphic-object-excel.md)** object that represents the picture for the right section of the footer. Used to set attributes of the picture.


## Syntax

 _expression_ . **RightFooterPicture**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

The  **RightFooterPicture** property itself is read-only, but not all of its properties are read-only.


## Example

The following example adds a picture titled Sample.jpg from the C: drive to the right section of the footer. This example assumes that a file called Sample.jpg exists on the C: drive.


```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.RightFooterPicture 
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
 
 ' Enable the image to show up in the right footer. 
 ActiveSheet.PageSetup.RightFooter = "&;G" 
 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

