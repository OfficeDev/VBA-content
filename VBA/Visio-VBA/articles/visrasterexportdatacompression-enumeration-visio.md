---
title: VisRasterExportDataCompression Enumeration (Visio)
keywords: vis_sdr.chm70750
f1_keywords:
- vis_sdr.chm70750
ms.prod: visio
api_name:
- Visio.VisRasterExportDataCompression
ms.assetid: 28fbd635-2b8f-7830-27d7-3b85c3879f19
ms.date: 06/08/2017
---


# VisRasterExportDataCompression Enumeration (Visio)



Specifies constants that identify the types of raster export data compression; passed to and returned by the ApplicationSettings.RasterExportDataCompression property.


|**Name**|**Value**|**Description**|**Applicable File Format**|
|:-----|:-----|:-----|:-----|
| **visRasterNone**|0|No compression; the default for BMP.|BMP|
| **visRasterRLE**|1|RLE compression.|BMP|
| **visRasterGroup3**|2|Group 3 compression.|TIFF|
| **visRasterPackbits**|3|Packbits compression.|TIFF|
| **visRasterGroup4**|4|Group 4 compression.|TIFF|
| **visRasterLZW**|5|LZW compression.|TIFF|
| **visRasterModifiedHuffman**|6|Modified Huffman compression.|TIFF|

