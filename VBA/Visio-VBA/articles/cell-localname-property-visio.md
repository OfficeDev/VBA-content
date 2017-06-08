---
title: Cell.LocalName Property (Visio)
keywords: vis_sdr.chm10113860
f1_keywords:
- vis_sdr.chm10113860
ms.prod: visio
api_name:
- Visio.Cell.LocalName
ms.assetid: 596bf196-6bbc-32f0-e508-03cdf4969a7f
ms.date: 06/08/2017
---


# Cell.LocalName Property (Visio)

Returns the local name of a cell. Read-only.


## Syntax

 _expression_ . **LocalName**

 _expression_ A variable that represents a **Cell** object.


### Return Value

String


## Remarks

A cell has both a local name and a universal name. The local name differs according to the locale for which Microsoft Windows is installed on the user's system. The universal name is the same regardless of locale.

To get the universal name of a cell, use the  **Name** property.


