---
title: Frame.PictureAlignment Property (Outlook Forms Script)
keywords: olfm10.chm2001700
f1_keywords:
- olfm10.chm2001700
ms.prod: outlook
ms.assetid: dda560cb-e002-1ae9-342a-ae2146bd3194
ms.date: 06/08/2017
---


# Frame.PictureAlignment Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the location of a background picture. Read/write.


## Syntax

 _expression_. **PictureAlignment**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The settings for  **PictureAlignment** are:



|**Value**|**Description**|
|:-----|:-----|
|0|The top left corner.|
|1|The top right corner.|
|2|The center.|
|3|The bottom left corner.|
|4|The bottom right corner.|
The  **PictureAlignment** property identifies which corner of the picture is the same as the corresponding corner of the control or container where the picture is used.

For example, setting  **PictureAlignment** to 0 means that the top left corner of the picture coincides with the top left corner of the control or container. Setting **PictureAlignment** to 2 positions the picture in the middle, relative to the height as well as the width of the control or container.

If you tile an image on a control or container, the setting of  **PictureAlignment** affects the tiling pattern. For example, if **PictureAlignment** is set to 0, the first copy of the image is laid in the upper left corner of the control or container and additional copies are tiled from left to right across each row. If **PictureAlignment** **PictureAlignment** is 2, the first copy of the image is laid at the center of the control or container, additional copies are laid to the left and right to complete the row, and additional rows are added to fill the control or container.

Setting the  **[PictureSizeMode](frame-picturesizemode-property-outlook-forms-script.md)** property to 2 overrides **PictureAlignment**. When  **PictureSizeMode** is set to 2, the picture fills the entire control or container.


