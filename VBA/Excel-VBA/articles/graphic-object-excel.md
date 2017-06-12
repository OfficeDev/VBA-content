---
title: Graphic Object (Excel)
keywords: vbaxl10.chm693072
f1_keywords:
- vbaxl10.chm693072
ms.prod: excel
api_name:
- Excel.Graphic
ms.assetid: 0ccdfb0d-effb-9fa4-8de9-b90688693375
ms.date: 06/08/2017
---


# Graphic Object (Excel)

Contains properties that apply to header and footer picture objects.


## Remarks

There are several properties of the  **[PageSetup](pagesetup-object-excel.md)** object that return the **Graphic** object.

Use the  **[CenterFooterPicture](pagesetup-centerfooterpicture-property-excel.md)** , **[CenterHeaderPicture](pagesetup-centerheaderpicture-property-excel.md)** , **[LeftFooterPicture](pagesetup-leftfooterpicture-property-excel.md)** , **[LeftHeaderPicture](pagesetup-leftheaderpicture-property-excel.md)** , **[RightFooterPicture](pagesetup-rightfooterpicture-property-excel.md)** , or **[RightHeaderPicture](pagesetup-rightheaderpicture-property-excel.md)** properties to return a **Graphic** object.


 **Note**  It is required that "&;G" is a part of the  **LeftFooter** string in order for the image to show up in the left footer.


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


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


