---
title: Document.GestureFormatSheet Property (Visio)
keywords: vis_sdr.chm10513605
f1_keywords:
- vis_sdr.chm10513605
ms.prod: visio
api_name:
- Visio.Document.GestureFormatSheet
ms.assetid: 26d3c27f-31ff-198c-5b40-8dc8b30b6362
ms.date: 06/08/2017
---


# Document.GestureFormatSheet Property (Visio)

Returns a reference to a document's Gesture Format sheet, which contains the line, fill, and text formatting that is applied to shapes drawn on the page. Read-only.


## Syntax

 _expression_ . **GestureFormatSheet**

 _expression_ A variable that represents a **Document** object.


### Return Value

Shape


## Remarks

By default, a new shape inherits all its formatting from the document's default styles. However, if the Gesture Format sheet contains local formatting, that formatting is applied to the new shape. Use the  **FillStyle** , **LineStyle** , and **TextStyle** properties to apply local formatting to the Gesture Format **Shape** object.

Gesture Format sheet formatting does not apply to instances of masters, connectors, pasted objects, or embedded objects.

A document's Gesture Format sheet is cleared automatically when a document is opened.

If a user attempts to use formatting commands on the menus and toolbars to make changes to any shape, but no shapes are currently selected, this formatting is stored in the gesture format sheet and applied to new shapes the user draws.


