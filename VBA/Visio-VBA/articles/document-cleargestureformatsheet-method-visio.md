---
title: Document.ClearGestureFormatSheet Method (Visio)
keywords: vis_sdr.chm10516120
f1_keywords:
- vis_sdr.chm10516120
ms.prod: visio
api_name:
- Visio.Document.ClearGestureFormatSheet
ms.assetid: 46f411b0-b822-cc7c-62e3-0b756d211d5d
ms.date: 06/08/2017
---


# Document.ClearGestureFormatSheet Method (Visio)

Clears local formatting in a document's Gesture Format sheet.


## Syntax

 _expression_ . **ClearGestureFormatSheet**

 _expression_ A variable that represents a **Document** object.


### Return Value

Nothing


## Remarks

Any shapes drawn after the Gesture Format sheet is cleared inherit their line, fill, and text formatting from the document's default styles.

A document's Gesture Format sheet also gets cleared automatically when the document is opened.

For details about the Gesture Format sheet, see the  **GestureFormatSheet** property.


