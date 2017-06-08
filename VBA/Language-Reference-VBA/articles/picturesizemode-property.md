---
title: PictureSizeMode Property
keywords: fm20.chm2001730
f1_keywords:
- fm20.chm2001730
ms.prod: office
api_name:
- Office.PictureSizeMode
ms.assetid: bb186d64-4e21-4ab5-3949-430c737e733d
ms.date: 06/08/2017
---


# PictureSizeMode Property



Specifies how to display the background picture on a control, form, or page.
 **Syntax**
 _object_. **PictureSizeMode** [= _fmPictureSizeMode_ ]
The  **PictureSizeMode** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmPictureSizeMode_|Optional. The action to take if the picture and the form or page that contains it are not the same size.|
 **Settings**
The settings for  _fmPictureSizeMode_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmPictureSizeModeClip_|0|Crops any part of the picture that is larger than the form or page (default).|
| _fmPictureSizeModeStretch_|1|Stretches the picture to fill the form or page. This setting distorts the picture in either the horizontal or vertical direction.|
| _fmPictureSizeModeZoom_|3|Enlarges the picture, but does not distort the picture in either the horizontal or vertical direction.|
 **Remarks**
The  **fmPictureSizeModeClip** setting indicates you want to show the picture in its original size and scale. If the form or page is smaller than the picture, this setting only shows the part of the picture that fits within the form or page.
The  **fmPictureSizeModeStretch** and **fmPictureSizeModeZoom** settings both enlarge the image, but **fmPictureSizeModeStretch** causes distortion. The **fmPictureSizeModeStretch** setting enlarges the image horizontally and vertically until the image reaches the corresponding edges of the[container](vbe-glossary.md) or control. The **fmPictureSizeModeZoom** setting enlarges the image until it reaches either the horizontal or vertical edges of the container or control. If the image reaches the horizontal edges first, any remaining distance to the vertical edges remains blank. If it reaches the vertical edges first, any remaining distance to the horizontal edges remains blank.

