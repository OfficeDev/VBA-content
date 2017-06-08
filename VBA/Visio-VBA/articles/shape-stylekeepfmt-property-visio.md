---
title: Shape.StyleKeepFmt Property (Visio)
keywords: vis_sdr.chm11214450
f1_keywords:
- vis_sdr.chm11214450
ms.prod: visio
api_name:
- Visio.Shape.StyleKeepFmt
ms.assetid: 22403064-fa5d-c0cf-8ee7-0f8ae2895d3c
ms.date: 06/08/2017
---


# Shape.StyleKeepFmt Property (Visio)

Applies a style to an object while preserving local formatting. Read/write.


## Syntax

 _expression_ . **StyleKeepFmt**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

Beginning with Microsoft Visio 2002, setting  **StyleKeepFmt** to an empty string ("") will cause the master's style to be reapplied to the shape. (Earlier versions generate a "no such style" exception.) If the shape has no master, its style remains unchanged.


