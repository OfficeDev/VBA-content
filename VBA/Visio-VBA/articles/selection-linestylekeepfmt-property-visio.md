---
title: Selection.LineStyleKeepFmt Property (Visio)
keywords: vis_sdr.chm11113850
f1_keywords:
- vis_sdr.chm11113850
ms.prod: visio
api_name:
- Visio.Selection.LineStyleKeepFmt
ms.assetid: 63703d4e-34b6-9b53-c2c1-b7503d0c3986
ms.date: 06/08/2017
---


# Selection.LineStyleKeepFmt Property (Visio)

Applies a line style to an object while preserving local formatting. Read/write.


## Syntax

 _expression_ . **LineStyleKeepFmt**

 _expression_ A variable that represents a **Selection** object.


### Return Value

String


## Remarks

Setting a style to a nonexistent style generates an error. Setting one kind of style to an existing style of another kind (for example, setting the  **LineStyleKeepFmt** property to a fill style) does nothing. Setting one kind of style to an existing style that has more than one set of attributes changes only the attributes for that component (for example, setting the **LineStyleKeepFmt** property to a style that has line, text, and fill attributes changes only the line attributes).

Beginning with Microsoft Visio 2002, setting  **LineStyleKeepFmt** to a zero-length string ("") will cause the master's style to be reapplied to the selection or shape. (Earlier versions generate a "no such style" exception.) If the selection or shape has no master, its style remains unchanged.


