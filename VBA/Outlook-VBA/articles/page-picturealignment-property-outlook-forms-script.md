---
title: Page.PictureAlignment Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c52f0b5b-c703-d9d6-1bae-e4fe9b696cf8
ms.date: 06/08/2017
---


# Page.PictureAlignment Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the location of a background picture. Read/write.


## Syntax

 _expression_. **PictureAlignment**

 _expression_A variable that represents a  **Page** object.


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

Setting the  **[PictureSizeMode](page-picturesizemode-property-outlook-forms-script.md)** property to 2 overrides **PictureAlignment**. When  **PictureSizeMode** is set to 2, the picture fills the entire control or container.


