---
title: Document.DefaultStyle Property (Visio)
keywords: vis_sdr.chm10513390
f1_keywords:
- vis_sdr.chm10513390
ms.prod: visio
api_name:
- Visio.Document.DefaultStyle
ms.assetid: e8fb078f-72cd-b4ae-1685-c0c02a265d3e
ms.date: 06/08/2017
---


# Document.DefaultStyle Property (Visio)

Gets the default fill style of a document or sets the default fill, line, and text styles of a document. Read/write.


## Syntax

 _expression_ . **DefaultStyle**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

A document's  **DefaultStyle** property returns the same value as its **DefaultFillStyle** property. Setting the **DefaultStyle** property is equivalent to setting the **DefaultFillStyle** , **DefaultLineStyle** , and **DefaultTextStyle** properties individually to the same multiple-attribute style. The fill, line, and text attributes of the document's default style are applied to new shapes created with the Microsoft Visio drawing tools or with the **Draw** methods by Automation.


