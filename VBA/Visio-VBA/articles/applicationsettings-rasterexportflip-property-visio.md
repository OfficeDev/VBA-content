---
title: ApplicationSettings.RasterExportFlip Property (Visio)
keywords: vis_sdr.chm16262550
f1_keywords:
- vis_sdr.chm16262550
ms.prod: visio
api_name:
- Visio.RasterExportFlip
ms.assetid: 1aa94fd4-7d2e-a2db-3291-c86ac4e22573
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportFlip Property (Visio)

Determines the flip that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.


## Syntax

 _expression_ . **RasterExportFlip**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **[VisRasterExportFlip](visrasterexportflip-enumeration-visio.md)**


## Remarks

The value of the  **RasterExportFlip** property must be either **visRasterNoFlip** or some combination (one, the other, or the sum of both) of the remaining two of the following **VisRasterExportFlip** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterNoFlip**|0|No flip, the default.|
| **visRasterFlipHorizontal**|1|Flip horizontally.|
| **visRasterFlipVertical**|2|Flip vertically.|
For any given session of Microsoft Visio, when the  **RasterExportFlip** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportFlip** property corresponds to the flip selected in the **BMP Output Options**,  **GIF Output Options**,  **JPG Output Options**,  **PNG Output Options**, or  **TIFF Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Windows Bitmap (*.bmp; *.dib)**,  **Graphics Interchange Format (*.gif)**,  **JPEG File Interchange Format (*.jpg)**,  **Portable Network Graphics (*.png)**, or  **Tag Image File Format (*.tif)**, and then click  **Save**.)


