---
title: ApplicationSettings.RasterExportQuality Property (Visio)
keywords: vis_sdr.chm16262570
f1_keywords:
- vis_sdr.chm16262570
ms.prod: visio
api_name:
- Visio.RasterExportQuality
ms.assetid: 6864bbfd-bb2d-721f-4146-f66974318929
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportQuality Property (Visio)

Determines the export quality that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a JPG file. Read/write.


## Syntax

 _expression_ . **RasterExportQuality**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **Long**


## Remarks

The default setting of the  **RasterExportQuality** property is 75%. The range of values is between 0% and 100%. Set the property value to a **Long** that corresponds to the quality percentage that you want.

For any given session of Microsoft Visio, when the  **RasterExportQuality** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportQuality** property corresponds to the **Quality** setting in the **JPG Output Options** dialog box in the Microsoft Visio user interface. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **JPEG File Interchange Format (*.jpg)**, and then click  **Save**.)


