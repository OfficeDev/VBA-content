---
title: ApplicationSettings.RasterExportDataCompression Property (Visio)
keywords: vis_sdr.chm16262525
f1_keywords:
- vis_sdr.chm16262525
ms.prod: visio
api_name:
- Visio.RasterExportDataCompression
ms.assetid: cec938db-1368-7c05-a264-b69ae334a249
ms.date: 06/08/2017
---


# ApplicationSettings.RasterExportDataCompression Property (Visio)

Determines the data compression algorithm that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP or TIFF file. Read/write.


## Syntax

 _expression_ . **RasterExportDataCompression**

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **[VisRasterExportDataCompression](visrasterexportdatacompression-enumeration-visio.md)**


## Remarks

The value of the  **RasterExportDataCompression** property must be one of the following **VisRasterExportDataCompression** constants.



|**Constant**|**Value**|**Description**|**Applicable File Format**|
|:-----|:-----|:-----|:-----|
| **visRasterNone**|0|No compression, the default for BMP.|BMP|
| **visRasterRLE**|1|RLE compression.|BMP|
| **visRasterGroup3**|2|Group 3 compression.|TIFF|
| **visRasterPackbits**|3|Packbits compression.|TIFF|
| **visRasterGroup4**|4|Group 4 compression.|TIFF|
| **visRasterLZW**|5|LZW compression.|TIFF|
| **visRasterModifiedHuffman**|6|Modified Huffman compression.|TIFF|
For any given session of Microsoft Visio, when the  **RasterExportDataCompression** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportDataCompression** property corresponds to the **Data compression** setting in the **BMP Output Options** or **TIFF Output Options** dialog box. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **Windows Bitmap (*.bmp; *.dib)** or **Tag Image File Format (*.tif)**, and then click  **Save**.)


