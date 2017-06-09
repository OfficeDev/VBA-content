---
title: Window.ShowPageOutline Property (Visio)
keywords: vis_sdr.chm11651615
f1_keywords:
- vis_sdr.chm11651615
ms.prod: visio
api_name:
- Visio.Window.ShowPageOutline
ms.assetid: 0e1f0413-1619-0e4f-ad44-e810ee2a38d1
ms.date: 06/08/2017
---


# Window.ShowPageOutline Property (Visio)

Determines whether the drawing page outline is displayed in the Microsoft Visio drawing window. Read/write.


## Syntax

 _expression_ . **ShowPageOutline**

 _expression_ A variable that represents a **Window** object.


### Return Value

Boolean


## Remarks

The default value is  **True** (the page outline is displayed), which is also the default Visio behavior. You can use the **ShowPageOutline** property to prevent display of the page outline in any Visio drawing window, including page, master, and group windows. Attempting to set **ShowPageOutline** for other windows, including stencil windows, ShapeSheet windows, and icon windows, will throw an exception.

Setting  **ShowPageOutline** to **False** does not hide the page grid. To hide the grid, use the **Window.ShowGrid** property.

The  **ShowPageOutline** property setting is valid only for a given window at run time, and is not persisted (saved) in either the Visio document or the Windows registry.


