---
title: VisRunTypes Enumeration (Visio)
keywords: vis_sdr.chm70095
f1_keywords:
- vis_sdr.chm70095
ms.prod: visio
ms.assetid: ec2d9be9-2e7a-50a1-b589-48c30a68b424
ms.date: 06/08/2017
---


# VisRunTypes Enumeration (Visio)

Run-type constants to be passed to the  **Characters.RunBegin** and **Characters.RunEnd** properties.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visCharPropRow**|1|Reports runs of characters that have common character properties. Corresponds to a set of characters covered by one row in a shape's Character section.|
| **visFieldRun**|20|Reports runs whose boundaries are between characters that are and aren't the result of the expansion of a text field, or between characters that are the result of the expansion of distinct text fields.|
| **visParaPropRow**|2|Reports runs of characters that have common paragraph properties. Corresponds to a set of characters covered by one row in the shape's Paragraph section.|
| **visParaRun**|11|Reports runs whose boundaries are between successive paragraphs in the shape's text. Mimics triple-clicking to select text.|
| **visTabPropRow**|3|Reports runs of characters that have common tab properties. Corresponds to a set of characters that are covered by one row in shape's Tabs section.|
| **visWordRun**|10|Reports runs whose boundaries are between successive words in a shape's text. Mimics double-clicking to select text.|

