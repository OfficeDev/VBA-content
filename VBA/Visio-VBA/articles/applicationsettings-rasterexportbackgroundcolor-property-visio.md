---
title: ApplicationSettings.RasterExportBackgroundColor Property (Visio)
keywords: vis_sdr.chm16262555
f1_keywords:
- vis_sdr.chm16262555
ms.prod: visio
api_name:
- Visio.RasterExportBackgroundColor
ms.assetid: 25591439-b332-af75-dec0-562cd261a453
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportBackgroundColor Property (Visio)

Determines the background color that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.


## Syntax

 _expression_ . **RasterExportBackgroundColor**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **OLE_COLOR**


## Remarks

The value of the  **RasterExportBackgroundColor** property must be a valid **OLE_COLOR** color. The default setting for a given session of Microsoft Visio is **RGB(255, 255,255)** , which means that Visio applies a white background color on export.

For any session of Visio, when the  **RasterExportBackgroundColor** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportBackgroundColor** property corresponds to the background color selected in the **BMP Output Options**,  **GIF Output Options**,  **JPG Output Options**,  **PNG Output Options**, or  **TIFF Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Windows Bitmap (*.bmp; *.dib)**,  **Graphics Interchange Format (*.gif)**,  **JPEG File Interchange Format (*.jpg)**,  **Portable Network Graphics (*.png)**, or  **Tag Image File Format (*.tif)**, and then click  **Save**.)


