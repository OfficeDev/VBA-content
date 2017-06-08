---
title: ApplicationSettings.RasterExportTransparencyColor Property (Visio)
keywords: vis_sdr.chm16262560
f1_keywords:
- vis_sdr.chm16262560
ms.prod: visio
api_name:
- Visio.RasterExportTransparencyColor
ms.assetid: 39806af2-1bdd-d659-134f-9cd86110e195
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportTransparencyColor Property (Visio)

Determines the transparency color that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a GIF or PNG file. Read/write.


## Syntax

 _expression_ . **RasterExportTransparencyColor**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **OLE_COLOR**


## Remarks

Microsoft Visio ignores this setting if  **RasterExportUseTransparencyColor** is **False**.

The value of the  **RasterExportTransparencyColor** property must be a valid **OLE_COLOR** color. The default is **RGB(255,255,255)** , which means that Visio applies a white transparency color on export.

For any given session of Microsoft Visio, when the  **RasterExportTransparencyColor** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportTransparencyColor** property corresponds to the transparency color selected in the **GIF Output Options** or **PNG Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Graphics Interchange Format (*.gif)** or **Portable Network Graphics (*.png)**, and then click  **Save**.)


