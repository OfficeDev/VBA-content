---
title: Page.PictureTiling Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1a8e3ed1-a71e-dbdb-8adc-52d6b7eb0977
ms.date: 06/08/2017
---


# Page.PictureTiling Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a picture is repeated across the background of the object. Read/write.


## Syntax

 _expression_. **PictureTiling**

 _expression_A variable that represents a  **Page** object.


## Remarks

 **True** if the picture is tiled across the background, **False** otherwise (default).

If a picture is smaller than the form or page that contains it, you can tile the picture on the form or page.

The tiling pattern depends on the current setting of the  **[PictureAlignment](page-picturealignment-property-outlook-forms-script.md)** and **[PictureSizeMode](page-picturesizemode-property-outlook-forms-script.md)** properties. For example, if **PictureAlignment** is set to 0, the tiling pattern starts at the upper left and repeats the picture across the form or page and down the height of the form or page. If **PictureSizeMode** is set to 0, the tiling pattern crops the last tile if it doesn't completely fit on the form or page.


