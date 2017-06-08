---
title: Shape.Style Property (Visio)
keywords: vis_sdr.chm11251160
f1_keywords:
- vis_sdr.chm11251160
ms.prod: visio
api_name:
- Visio.Shape.Style
ms.assetid: beba03ba-6926-d2db-4e36-652d05c2925c
ms.date: 06/08/2017
---


# Shape.Style Property (Visio)

Gets or sets the style for a  **Shape** object. Read/write.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

If a style consists of different text, line, and fill styles, the  **Style** property returns the fill style. If you set the **Style** property to a nonexistent style, your program generates an error.

To preserve local formatting, use the  **StyleKeepFmt** property.

Beginning with Visio 2002, setting  **Style** to an empty string ("") will cause the master's style to be reapplied to the shape. (Earlier versions generate a "no such style" exception.) If the shape has no master, its style remains unchanged.


