---
title: Window.MergePosition Property (Visio)
keywords: vis_sdr.chm11650725
f1_keywords:
- vis_sdr.chm11650725
ms.prod: visio
api_name:
- Visio.Window.MergePosition
ms.assetid: 0856bcec-191d-5c9c-f44a-cd430bc3ceb8
ms.date: 06/08/2017
---


# Window.MergePosition Property (Visio)

Specifies the left-to-right tab position of a merged anchored window. Read/write.


## Syntax

 _expression_ . **MergePosition**

 _expression_ A variable that represents a **Window** object.


### Return Value

Long


## Remarks

If there are  _n_ tabs, the leftmost position is 1 and the rightmost position is _n_ . If the windows are merged in a docked stencil fashion, 1 is the topmost and _n_ is the bottommost. A value of zero (0) means that the window is not merged.

The  **MergePosition** property applies only to anchored windows. If the **Window** object is an MDI frame window, Microsoft Visio raises an exception.

Use the  **Type** property to determine window type.


