---
title: PageSetup.PaperSize Property (Excel)
keywords: vbaxl10.chm473091
f1_keywords:
- vbaxl10.chm473091
ms.prod: excel
api_name:
- Excel.PageSetup.PaperSize
ms.assetid: 7c26e996-8399-31b4-8e53-772de8bf8eb2
ms.date: 06/08/2017
---


# PageSetup.PaperSize Property (Excel)

Returns or sets the size of the paper. Read/write [XlPaperSize](xlpapersize-enumeration-excel.md).


## Syntax

 _expression_ . **PaperSize**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks



| **XlPaperSize** can be one of these **XlPaperSize** constants.|
| **xlPaper11x17** . 11 in. x 17 in.|
| **xlPaperA4** . A4 (210 mm x 297 mm)|
| **xlPaperA5** . A5 (148 mm x 210 mm)|
| **xlPaperB5** . A5 (148 mm x 210 mm)|
| **xlPaperDsheet** . D size sheet|
| **xlPaperEnvelope11** . Envelope #11 (4-1/2 in. x 10-3/8 in.)|
| **xlPaperEnvelope14** . Envelope #14 (5 in. x 11-1/2 in.)|
| **xlPaperEnvelopeB4** . Envelope B4 (250 mm x 353 mm)|
| **xlPaperEnvelopeB6** . Envelope B6 (176 mm x 125 mm)|
| **xlPaperEnvelopeC4** . Envelope C4 (229 mm x 324 mm)|
| **xlPaperEnvelopeC6** . Envelope C6 (114 mm x 162 mm)|
| **xlPaperEnvelopeDL** . Envelope DL (110 mm x 220 mm)|
| **xlPaperEnvelopeMonarch** . Envelope Monarch (3-7/8 in. x 7-1/2 in.)|
| **xlPaperEsheet** . E size sheet|
| **xlPaperFanfoldLegalGerman** . German Legal Fanfold (8-1/2 in. x 13 in.)|
| **xlPaperFanfoldUS** . U.S. Standard Fanfold (14-7/8 in. x 11 in.)|
| **xlPaperLedger** . Ledger (17 in. x 11 in.)|
| **xlPaperLetter** . Letter (8-1/2 in. x 11 in.)|
| **xlPaperNote** . Note (8-1/2 in. x 11 in.)|
| **xlPaperStatement** . Statement (5-1/2 in. x 8-1/2 in.)|
| **xlPaperUser** . User-defined|
| **xlPaper10x14** . 10 in. x 14 in.|
| **xlPaperA3** . A3 (297 mm x 420 mm)|
| **xlPaperA4Small** . A4 Small (210 mm x 297 mm)|
| **xlPaperB4** . B4 (250 mm x 354 mm)|
| **xlPaperCsheet** . C size sheet|
| **xlPaperEnvelope10** . Envelope #10 (4-1/8 in. x 9-1/2 in.)|
| **xlPaperEnvelope12** . Envelope #12 (4-1/2 in. x 11 in.)|
| **xlPaperEnvelope9** . Envelope #9 (3-7/8 in. x 8-7/8 in.)|
| **xlPaperEnvelopeB5** . Envelope B5 (176 mm x 250 mm)|
| **xlPaperEnvelopeC3** . Envelope C3 (324 mm x 458 mm)|
| **xlPaperEnvelopeC5** . Envelope C5 (162 mm x 229 mm)|
| **xlPaperEnvelopeC65** . Envelope C65 (114 mm x 229 mm)|
| **xlPaperEnvelopeItaly** . Envelope (110 mm x 230 mm)|
| **xlPaperEnvelopePersonal** . Envelope (3-5/8 in. x 6-1/2 in.)|
| **xlPaperExecutive** . Executive (7-1/2 in. x 10-1/2 in.)|
| **xlPaperFanfoldStdGerman** . German Legal Fanfold (8-1/2 in. x 13 in.)|
| **xlPaperFolio** . Folio (8-1/2 in. x 13 in.)|
| **xlPaperLegal** . Legal (8-1/2 in. x 14 in.)|
| **xlPaperLetterSmall** . Letter Small (8-1/2 in. x 11 in.)|
| **xlPaperQuarto** . Quarto (215 mm x 275 mm)|
| **xlPaperTabloid** . Tabloid (11 in. x 17 in.)|

 **Note**  Some printers may not support all of these paper sizes.


## Example

This example sets the paper size to legal for Sheet1.


```vb
Worksheets("Sheet1").PageSetup.PaperSize = xlPaperLegal
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

