---
title: ApplicationSettings.RasterExportColorFormat Property (Visio)
keywords: vis_sdr.chm16262535
f1_keywords:
- vis_sdr.chm16262535
ms.prod: visio
api_name:
- Visio.RasterExportColorFormat
ms.assetid: 8306b2c1-d0a0-41ae-16de-0deb4d881604
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportColorFormat Property (Visio)

Determines the color format that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, JPG, PNG, or TIFF file. Read/write.


## Syntax

 _expression_ . **RasterExportColorFormat**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **[VisRasterExportColorFormat](visrasterexportcolorformat-enumeration-visio.md)**


## Remarks

The value of the  **RasterExportColorFormat** property must be one of the following **VisRasterExportColorFormat** constants.



|**Constant**|**Value**|**Description**|**Applicable File Formats**|
|:-----|:-----|:-----|:-----|
| **visRasterBiLevel**|0|Bi-level color format|BMP, PNG, TIFF|
| **visRaster16Color**|1|16-color format|BMP, PNG, TIFF|
| **visRaster256Color**|2|256-color format|BMP, PNG, TIFF|
| **visRaster24Bit**|3|24-bit color format; the default for PNG, TIFF, and BMP|BMP, PNG, TIFF|
| **visRasterRGB**|4|RGB color format; the default for JPG|JPG|
| **visRasterYCC**|5|YCC color format|JPG|
| **visRasterGrayScale**|6|Grayscale color format|JPG|
| **visRasterCMYK**|7|CMYK color format|JPG|
| **visRaster16ColorGrayScale**|8|16-color grayscale color format|TIFF|
| **visRaster256ColorGrayScale**|9|256-color grayscale color format|TIFF|
| **visRaster16Bit**|10|16-bit color format|BMP|
For any given session of Microsoft Visio, when the  **RasterExportColorFormat** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportColorFormat** property corresponds to the **Color format** setting in the **BMP Output Options**,  **JPG Output Options**,  **PNG Output Options**, or  **TIFF Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Windows Bitmap (*.bmp; *.dib)**,  **JPEG File Interchange Format (*.jpg)**,  **Portable Network Graphics (*.png)**, or  **Tag Image File Format (*.tif)**, and then click  **Save**.)


