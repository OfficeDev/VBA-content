---
title: PictureTiling Property
keywords: fm20.chm2001740
f1_keywords:
- fm20.chm2001740
ms.prod: office
api_name:
- Office.PictureTiling
ms.assetid: 9bf2a163-0454-b959-0261-b2a9fd7f6bfa
ms.date: 06/08/2017
---


# PictureTiling Property



Lets you tile a picture in a form or page.
 **Syntax**
 _object_. **PictureTiling** [= _Boolean_ ]
The  **PictureTiling** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether a picture is repeated across a background.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The picture is tiled across the background.|
|**False**|The picture is not tiled across the background (default).|
 **Remarks**
If a picture is smaller than the form or page that contains it, you can tile the picture on the form or page.
The tiling pattern depends on the current setting of the  **PictureAlignment** and **PictureSizeMode** properties. For example, if **PictureAlignment** is set to **fmPictureAlignmentTopLeft**, the tiling pattern starts at the upper left and repeats the picture across the form or page and down the height of the form or page. If **PictureSizeMode** is set to **fmPictureSizeModeClip**, the tiling pattern crops the last tile if it doesn't completely fit on the form or page.

