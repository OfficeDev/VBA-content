---
title: Selection.Style Property (Visio)
keywords: vis_sdr.chm11151130
f1_keywords:
- vis_sdr.chm11151130
ms.prod: visio
api_name:
- Visio.Selection.Style
ms.assetid: f0853c43-14b4-bcd9-eb07-fbc0312e106b
ms.date: 06/08/2017
---


# Selection.Style Property (Visio)

Gets or sets the style for a  **Selection** object. Read/write.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Selection** object.


### Return Value

String


## Remarks

If a style consists of different text, line, and fill styles, the  **Style** property returns the fill style. If you set the **Style** property to a nonexistent style, your program generates an error.

To preserve local formatting, use the  **StyleKeepFmt** property.

Beginning with Visio 2002, setting the  **Style** propery to an empty string ("") will cause the master's style to be reapplied to the shape. (Earlier versions generate a "no such style" exception.) If the shape has no master, its style remains unchanged.


