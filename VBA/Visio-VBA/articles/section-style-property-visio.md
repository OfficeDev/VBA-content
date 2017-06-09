---
title: Section.Style Property (Visio)
keywords: vis_sdr.chm15751155
f1_keywords:
- vis_sdr.chm15751155
ms.prod: visio
api_name:
- Visio.Section.Style
ms.assetid: cd8d041d-126e-7983-0a13-48fb9f5f5df6
ms.date: 06/08/2017
---


# Section.Style Property (Visio)

Gets the style that contains a  **Section** object. Read-only.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Section** object.


### Return Value

Style


## Remarks

If a style consists of different text, line, and fill styles, the  **Style** property returns the fill style.

If a  **Section** object is in a style, its **Style** property returns the style that contains the cell, and its **Shape** property returns **Nothing** .

If a  **Section** object is in a shape, its **Shape** property returns the shape that contains the cell, and its **Style** property returns **Nothing** .


