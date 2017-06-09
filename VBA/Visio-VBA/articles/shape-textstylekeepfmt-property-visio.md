---
title: Shape.TextStyleKeepFmt Property (Visio)
keywords: vis_sdr.chm11214535
f1_keywords:
- vis_sdr.chm11214535
ms.prod: visio
api_name:
- Visio.Shape.TextStyleKeepFmt
ms.assetid: add41319-8b81-a803-46e2-697df37eb731
ms.date: 06/08/2017
---


# Shape.TextStyleKeepFmt Property (Visio)

Applies a text style to an object while preserving local formatting. Read/write.


## Syntax

 _expression_ . **TextStyleKeepFmt**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

Setting a style to a nonexistent style generates an error. Setting one kind of style to an existing style of another kind (for example, setting the  **TextStyleKeepFmt** property to a fill style) does nothing. Setting one kind of style to an existing style that has more than one set of attributes changes only the attributes for that component (for example, setting the **TextStyleKeepFmt** property to a style that has line, text, and fill attributes changes only the text attributes).

Beginning with Visio 2002, setting  **TextStyleKeepFmt** to an empty string ("") will cause the master's style to be reapplied to the selection or shape. (Earlier versions generate a "no such style" exception.) If the selection or shape has no master, its style remains unchanged.


