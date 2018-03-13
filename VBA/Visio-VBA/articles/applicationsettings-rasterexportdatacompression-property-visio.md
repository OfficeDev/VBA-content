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



| <strong>Constant</strong>                 | <strong>Value</strong> | <strong>Description</strong>         | <strong>Applicable File Format</strong> |
|:------------------------------------------|:-----------------------|:-------------------------------------|:----------------------------------------|
| <strong>visRasterNone</strong>            | 0                      | No compression, the default for BMP. | BMP                                     |
| <strong>visRasterRLE</strong>             | 1                      | RLE compression.                     | BMP                                     |
| <strong>visRasterGroup3</strong>          | 2                      | Group 3 compression.                 | TIFF                                    |
| <strong>visRasterPackbits</strong>        | 3                      | Packbits compression.                | TIFF                                    |
| <strong>visRasterGroup4</strong>          | 4                      | Group 4 compression.                 | TIFF                                    |
| <strong>visRasterLZW</strong>             | 5                      | LZW compression.                     | TIFF                                    |
| <strong>visRasterModifiedHuffman</strong> | 6                      | Modified Huffman compression.        | TIFF                                    |

For any given session of Microsoft Visio, when the  **RasterExportDataCompression** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  <strong>RasterExportDataCompression</strong> property corresponds to the <strong>Data compression</strong> setting in the <strong>BMP Output Options</strong> or <strong>TIFF Output Options</strong> dialog box. (Click the <strong>File</strong> tab, click <strong>Save As</strong>, in the  <strong>Save as type</strong> list, select <strong>Windows Bitmap (*.bmp; *.dib)</strong> or <strong>Tag Image File Format (*.tif)</strong>, and then click  <strong>Save</strong>.)


