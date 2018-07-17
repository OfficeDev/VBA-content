---
title: ApplicationSettings.RasterExportUseTransparencyColor Property (Visio)
keywords: vis_sdr.chm16262565
f1_keywords:
- vis_sdr.chm16262565
ms.prod: visio
api_name:
- Visio.RasterExportUseTransparencyColor
ms.assetid: 1fd93b1b-8b35-a82a-17f5-0fa2ffa819a7
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportUseTransparencyColor Property (Visio)

Determines whether Microsoft Visio applies, to the exported image, the transparency color that is specified in the  **RasterExportTransparencyColor** property when you call the **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a GIF or PNG file. Read/write.


## Syntax

 _expression_ . **RasterExportUseTransparencyColor**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **Boolean**


## Remarks

The default is  **False** , which means that Microsoft Visio does not apply the transparency color to the page on export.

For any given session of Microsoft Visio, when the  **RasterExportUseTransparencyColor** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportUseTransparencyColor** property corresponds to whether the **Transparency color** box in the **GIF Output Options** or the **PNG Output Options** dialog box is selected. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Graphics Interchange Format (*.gif)** or **Portable Network Graphics (*.png)**, and then click  **Save**.)


