---
title: Selection.StyleKeepFmt Property (Visio)
keywords: vis_sdr.chm11114450
f1_keywords:
- vis_sdr.chm11114450
ms.prod: visio
api_name:
- Visio.Selection.StyleKeepFmt
ms.assetid: b56bfda8-0076-0114-b231-bb7c649c6310
ms.date: 06/08/2017
---


# Selection.StyleKeepFmt Property (Visio)

Applies a style to an object while preserving local formatting. Read/write.


## Syntax

 _expression_ . **StyleKeepFmt**

 _expression_ A variable that represents a **Selection** object.


### Return Value

String


## Remarks

Setting a style to a nonexistent style generates an error.

Beginning with Microsoft Visio 2002, setting  **StyleKeepFmt** to an empty string ("") will cause the master's style to be reapplied to the shape. (Earlier versions generate a "no such style" exception.) If the shape has no master, its style remains unchanged.


