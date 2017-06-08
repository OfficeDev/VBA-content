---
title: Page.PrintTileCount Property (Visio)
keywords: vis_sdr.chm10950800
f1_keywords:
- vis_sdr.chm10950800
ms.prod: visio
api_name:
- Visio.Page.PrintTileCount
ms.assetid: f15eff27-1d20-7151-e773-1ab4de4161db
ms.date: 06/08/2017
---


# Page.PrintTileCount Property (Visio)

Returns the number of print tiles for a drawing page. Read-only.


## Syntax

 _expression_ . **PrintTileCount**

 _expression_ A variable that represents a **Page** object.


### Return Value

Long


## Remarks

When drawings span multiple physical printer pages, you can use the  **PrintTileCount** property to determine the number of print tiles there are for a Microsoft Visio drawing page. You can use the **PrintTileCount** property with the **PrintTile** method to identify and print selected tiles of an active drawing page.


