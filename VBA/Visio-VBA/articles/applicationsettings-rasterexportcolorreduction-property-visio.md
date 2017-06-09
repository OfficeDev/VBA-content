---
title: ApplicationSettings.RasterExportColorReduction Property (Visio)
keywords: vis_sdr.chm16262530
f1_keywords:
- vis_sdr.chm16262530
ms.prod: visio
api_name:
- Visio.RasterExportColorReduction
ms.assetid: 7897f3aa-d7d1-4dcc-d4f1-9c38771c0122
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportColorReduction Property (Visio)

Determines the color reduction that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, PNG, or TIFF file. Read/write.


## Syntax

 _expression_ . **RasterExportColorReduction**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **[VisRasterExportColorReduction](visrasterexportcolorreduction-enumeration-visio.md)**


## Remarks

When you apply color reduction, you reduce the number of colors present in an image in order to make the resulting file smaller in size.

The value of the  **RasterExportColorReduction** property can be one of the following **VisRasterExportColorReduction** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterAdaptive**|0|Adaptive color reduction; the default for GIF files|
| **visRasterDiffusion**|1|Diffusion color reduction|
| **visRasterHalftone**|2|Halftone color reduction|
For BMP, PNG, and TIFF files, the default is no color reduction. 

For any given session of Microsoft Visio, when the  **RasterExportColorReduction** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportColorReduction** property corresponds to the **Color reduction** setting in the **BMP Output Options**,  **GIF Output Options**,  **PNG Output Options**, or  **TIFF Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Windows Bitmap (*.bmp; *.dib)**,  **Graphics Interchange Format (*.gif)**,  **Portable Network Graphics (*.png)**, or  **Tag Image File Format (*.tif)**, and then click  **Save**.)


