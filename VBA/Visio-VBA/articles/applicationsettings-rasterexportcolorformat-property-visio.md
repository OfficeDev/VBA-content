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



| <strong>Constant</strong>                   | <strong>Value</strong> | <strong>Description</strong>                            | <strong>Applicable File Formats</strong> |
|:--------------------------------------------|:-----------------------|:--------------------------------------------------------|:-----------------------------------------|
| <strong>visRasterBiLevel</strong>           | 0                      | Bi-level color format                                   | BMP, PNG, TIFF                           |
| <strong>visRaster16Color</strong>           | 1                      | 16-color format                                         | BMP, PNG, TIFF                           |
| <strong>visRaster256Color</strong>          | 2                      | 256-color format                                        | BMP, PNG, TIFF                           |
| <strong>visRaster24Bit</strong>             | 3                      | 24-bit color format; the default for PNG, TIFF, and BMP | BMP, PNG, TIFF                           |
| <strong>visRasterRGB</strong>               | 4                      | RGB color format; the default for JPG                   | JPG                                      |
| <strong>visRasterYCC</strong>               | 5                      | YCC color format                                        | JPG                                      |
| <strong>visRasterGrayScale</strong>         | 6                      | Grayscale color format                                  | JPG                                      |
| <strong>visRasterCMYK</strong>              | 7                      | CMYK color format                                       | JPG                                      |
| <strong>visRaster16ColorGrayScale</strong>  | 8                      | 16-color grayscale color format                         | TIFF                                     |
| <strong>visRaster256ColorGrayScale</strong> | 9                      | 256-color grayscale color format                        | TIFF                                     |
| <strong>visRaster16Bit</strong>             | 10                     | 16-bit color format                                     | BMP                                      |

For any given session of Microsoft Visio, when the  **RasterExportColorFormat** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportColorFormat** property corresponds to the **Color format** setting in the **BMP Output Options**,  **JPG Output Options**,  **PNG Output Options**, or  **TIFF Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Windows Bitmap (*.bmp; *.dib)**,  **JPEG File Interchange Format (*.jpg)**,  **Portable Network Graphics (*.png)**, or  **Tag Image File Format (*.tif)**, and then click  **Save**.)


