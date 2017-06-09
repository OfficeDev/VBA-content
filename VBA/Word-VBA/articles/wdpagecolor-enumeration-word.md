---
title: WdPageColor Enumeration (Word)
ms.prod: word
ms.assetid: 99f557e5-48c9-65e3-0ab4-fed727fa5890
ms.date: 06/08/2017
---


# WdPageColor Enumeration (Word)

Constants that represent background page color in reading mode, returned by and passed to [View.PageColor](view-pagecolor-property-word.md).


## Members



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdPageColorInverse**|2|Inverse page color. Renders the document content in a manner that resembles high-contrast black, although not necessarily exactly so. Some figures are rendered in full color on a black background.|
| **wdPageColorNone**|0|No page color, the default. The page background is rendered in white. Any assigned page background colors are ignored.|
| **wdPageColorSepia**|1|Sepia page color, RGB (112, 66, 20) at 80% transparency. Makes no changes to the contents of the document.|

