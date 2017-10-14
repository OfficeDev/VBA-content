---
title: Cell.Style Property (Visio)
keywords: vis_sdr.chm10151145
f1_keywords:
- vis_sdr.chm10151145
ms.prod: visio
api_name:
- Visio.Cell.Style
ms.assetid: 12eec8c7-706a-488e-ad3a-326c9f628f5c
ms.date: 06/08/2017
---


# Cell.Style Property (Visio)

Gets the style that contains a  **Cell** object. Read-only.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Style


## Remarks

If a style consists of different text, line, and fill styles, the  **Style** property returns the fill style.

If a  **Cell** object is in a style, its **Style** property returns the style that contains the cell, and its **Shape** property returns **Nothing** .

If a  **Cell** object is in a shape, its **Shape** property returns the shape that contains the cell, and its **Style** property returns **Nothing** .


