---
title: Document.PaperSize Property (Visio)
keywords: vis_sdr.chm10514020
f1_keywords:
- vis_sdr.chm10514020
ms.prod: visio
api_name:
- Visio.Document.PaperSize
ms.assetid: a51b3d26-e79e-d572-055f-fc1c4a94096e
ms.date: 06/08/2017
---


# Document.PaperSize Property (Visio)

Gets or sets the paper size of a document. Read/write.


## Syntax

 _expression_ . **PaperSize**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisPaperSizes


## Remarks

Setting the  **PaperSize** property is the equivalent of choosing a printer paper size on the **Print Setup** tab of the **Page Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow). The value of **PaperSize** can be one of the following **VisPaperSizes** constants.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visPaperSizeUnknown**| 0| Not known|
| **visPaperSizeLetter**| 1| Letter 8 1/2 x 11 in|
| **visPaperSizeLegal**| 5| Legal 8 1/2 x 14 in|
| **visPaperSizeA3**| 8| A3 297 x 420 mm|
| **visPaperSizeA4**| 9| A4 210 x 297 mm|
| **visPaperSizeA5**| 11| A5 148 x 210 mm|
| **visPaperSizeB4**| 12| B4 (JIS) 250 x 354 mm|
| **visPaperSizeB5**| 13| B5 (JIS) 182 x 257 mm|
| **visPaperSizeFolio**| 14| Folio 8 1/2 x 13 in|
| **visPaperSizeNote**| 18| Note 8 1/2 x 11 in|
| **visPaperSizeSizeC**| 24| C size sheet 17 x 22 in.|
| **visPaperSizeSizeD**| 25| D size sheet 22 x 34 in.|
| **visPaperSizeSizeE**| 26| E size sheet 34 x 44 in.|
If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocument.PaperSize**
    

