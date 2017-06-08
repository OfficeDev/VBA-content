---
title: WdHorizontalLineWidthType Enumeration (Word)
ms.prod: word
api_name:
- Word.WdHorizontalLineWidthType
ms.assetid: c36b6de0-6963-c92d-5e95-45e72eb4d2c2
ms.date: 06/08/2017
---


# WdHorizontalLineWidthType Enumeration (Word)

Specifies how Word interprets the width (length) of the specified horizontal line.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdHorizontalLineFixedWidth**|-2|Microsoft Word interprets the width (length) of the specified horizontal line as a fixed value (in points). This is the default value for horizontal lines added with the  **AddHorizontalLine** method. Setting the **Width** property for the **InlineShape** object associated with a horizontal line sets the **WidthType** property to this value.|
| **wdHorizontalLinePercentWidth**|-1|Word interprets the width (length) of the specified horizontal line as a percentage of the screen width. This is the default value for horizontal lines added with the  **AddHorizontalLineStandard** method. Setting the **PercentWidth** property on a horizontal line sets the **WidthType** property to this value.|

