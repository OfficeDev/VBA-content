---
title: PictureAlignment Property
keywords: fm20.chm5225078
f1_keywords:
- fm20.chm5225078
ms.prod: office
api_name:
- Office.PictureAlignment
ms.assetid: 5d497e60-7106-6278-a5c0-06ef06d6177f
ms.date: 06/08/2017
---


# PictureAlignment Property



Specifies the location of a background picture.
 **Syntax**
 _object_. **PictureAlignment** [= _fmPictureAlignment_ ]
The  **PictureAlignment** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmPictureAlignment_|Optional. The position where the picture aligns with the control.|
 **Settings**
The settings for  _fmPictureAlignment_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmPictureAlignmentTopLeft_|0|The top left corner.|
| _fmPictureAlignmentTopRight_|1|The top right corner.|
| _fmPictureAlignmentCenter_|2|The center.|
| _fmPictureAlignmentBottomLeft_|3|The bottom left corner.|
| _fmPictureAlignmentBottomRight_|4|The bottom right corner.|
 **Remarks**
The  **PictureAlignment** property identifies which corner of the picture is the same as the corresponding corner of the control or[container](vbe-glossary.md) where the picture is used.
For example, setting  **PictureAlignment** to **fmPictureAlignmentTopLeft** means that the top left corner of the picture coincides with the top left corner of the control or container. Setting **PictureAlignment** to **fmPictureAlignmentCenter** positions the picture in the middle, relative to the height as well as the width of the control or container.
If you tile an image on a control or container, the setting of  **PIctureAlignment** affects the tiling pattern. For example, if **PictureAlignment** is set to **fmPictureAlignmentUpperLeft**, the first copy of the image is laid in the upper left corner of the control or container and additional copies are tiled from left to right across each row. If **PictureAlignment** is **fmPictureAlignmentCenter**, the first copy of the image is laid at the center of the control or container, additional copies are laid to the left and right to complete the row, and additional rows are added to fill the control or container.

 **Note**  Setting the  **PictureSizeMode** property to **fmSizeModeStretch** overrides **PictureAlignment**. When **PictureSizeMode** is set to **fmSizeModeStretch**, the picture fills the entire control or container.


