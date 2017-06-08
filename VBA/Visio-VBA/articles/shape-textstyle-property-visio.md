---
title: Shape.TextStyle Property (Visio)
keywords: vis_sdr.chm11214530
f1_keywords:
- vis_sdr.chm11214530
ms.prod: visio
api_name:
- Visio.Shape.TextStyle
ms.assetid: 9436ba1b-f792-aed6-3936-b2d88a6dd2ea
ms.date: 06/08/2017
---


# Shape.TextStyle Property (Visio)

Gets or sets the text style for an object. Read/write.


## Syntax

 _expression_ . **TextStyle**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

Setting a style to a nonexistent style generates an error. Setting one kind of style to an existing style of another kind (for example, setting the  **TextStyle** property to a fill style) does nothing. Setting one kind of style to an existing style that has more than one set of attributes changes only the attributes for that component (for example, setting the **TextStyle** property to a style that has line, text, and fill attributes changes only the text attributes).

To preserve a shape's local formatting, use the  **TextStyleKeepFmt** property.

Beginning with Visio 2002, setting  **TextStyle** to an empty string ("") will cause the master's style to be reapplied to the selection or shape. (Earlier versions generate a "no such style" exception.) If the selection or shape has no master, its style remains unchanged.


