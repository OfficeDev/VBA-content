---
title: ApplicationSettings.RasterExportRotation Property (Visio)
keywords: vis_sdr.chm16262545
f1_keywords:
- vis_sdr.chm16262545
ms.prod: visio
api_name:
- Visio.RasterExportRotation
ms.assetid: 660b22ff-11b6-bfaf-1949-18e5e9c57d64
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportRotation Property (Visio)

Determines the rotation that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.


## Syntax

 _expression_ . **RasterExportRotation**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **[VisRasterExportRotation](visrasterexportrotation-enumeration-visio.md)**


## Remarks

The value of the  **RasterExportRotation** property must be one of the following **VisRasterExportRotation** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterNoRotation**|0|No rotation, the default.|
| **visRasterRotateLeft**|1|Rotate left.|
| **visRasterRotateRight**|2|Rotate right.|
For any given session of Microsoft Visio, when the  **RasterExportRotation** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportRotation** property corresponds to the style of rotation selected in the **Rotation** list in the **Output Options** dialog box for the corresponding file type in the Microsoft Visio user interface. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select the file type, and then click **Save**.)


