---
title: Row.Style Property (Visio)
keywords: vis_sdr.chm15814445
f1_keywords:
- vis_sdr.chm15814445
ms.prod: visio
api_name:
- Visio.Row.Style
ms.assetid: d11fac30-0349-e202-a3db-fab9c65665a1
ms.date: 06/08/2017
---


# Row.Style Property (Visio)

Gets the style that contains a  **Row** object. Read-only.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Row** object.


### Return Value

Style


## Remarks

If a style consists of different text, line, and fill styles, the  **Style** property returns the fill style.

If a  **Row** object is in a style, its **Style** property returns the style that contains the cell, and its **Shape** property returns **Nothing** .

If a  **Row** object is in a shape, its **Shape** property returns the shape that contains the cell, and its **Style** property returns **Nothing** .


