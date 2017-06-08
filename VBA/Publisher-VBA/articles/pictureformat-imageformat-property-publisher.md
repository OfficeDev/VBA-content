---
title: PictureFormat.ImageFormat Property (Publisher)
keywords: vbapb10.chm3604761
f1_keywords:
- vbapb10.chm3604761
ms.prod: publisher
api_name:
- Publisher.PictureFormat.ImageFormat
ms.assetid: a5523a1e-4dbf-5cd7-ba73-2a5570865ee6
ms.date: 06/08/2017
---


# PictureFormat.ImageFormat Property (Publisher)

Returns a  **PbImageFormat** constant that represents the image format of a picture as determined by Microsoft Windows Graphics Device Interface (GDI+). Read-only.


## Syntax

 _expression_. **ImageFormat**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

PbImageFormat


## Remarks

The  **ImageFormat** property applies to the original picture, rather than the placeholder picture, if there is one.

The  **ImageFormat** property value can be one of the **[PbImageFormat](pbimageformat-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

The  **ImageFormat** property indicates the format of the picture after it has been imported into the Windows environment, rather than its original file format. If the picture's file format is not natively supported by the Windows operating system, the picture is converted to an analogous format that is natively supported. As a result, the **pbImageFormatCMYKJPEG**,  **pbImageFormatDIB**,  **pbImageFormatEMF**,  **pbImageFormatGIF**, and  **pbImageFormatPICT** constants will rarely, if ever, be returned. Consult the table below for specific file format conversions.



|**File format**|**Constant returned**|
|:-----|:-----|
|.bmp, .dib, .gif, .pict|pbImageFormatPNG|
|.emf, .eps, .epfs|pbImageFormatWMF|
|CMYK .jfif, .jpeg, .jpg|pbImageFormatJPEG|
Windows GDI+ is the portion of the Microsoft Windows XP operating system and the Microsoft Windows Server 2003 operating system that provides two-dimensional vector graphics, imaging, and typography.


## Example

The following example prints a list of the .jpg and .jpeg images present in the active publication.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .ImageFormat = pbImageFormatJPEG Then 
 Debug.Print .Filename 
 End If 
 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
Next pgLoop 

```


