---
title: ApplicationSettings.RasterExportDataFormat Property (Visio)
keywords: vis_sdr.chm16262520
f1_keywords:
- vis_sdr.chm16262520
ms.prod: visio
api_name:
- Visio.RasterExportDataFormat
ms.assetid: e07c3f2e-469e-33bc-cd6d-0261cf7ec267
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportDataFormat Property (Visio)

Determines whether the exported raster image is interlaced or non-interlaced when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a GIF or PNG file. Read/write.


## Syntax

 _expression_ . **RasterExportDataFormat**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **[VisRasterExportDataFormat](visrasterexportdataformat-enumeration-visio.md)**


## Remarks

The value of the  **RasterExportDataFormat** property must be one of the following **VisRasterExportDataFormat** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterInterlace**|0|Interlace format, the default.|
| **visRasterNonInterlace**|1|Non-interlace format.|
For any given session of Microsoft Visio, when the  **RasterExportDataFormat** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportDataFormat** property corresponds to the **Data format** setting in the **GIF Output Options** or **PNG Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Graphics Interchange Format (*.gif)** or **Portable Network Graphics (*.png)**, and then click  **Save**.)


