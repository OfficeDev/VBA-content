---
title: Window.SelectionForDragCopy Property (Visio)
keywords: vis_sdr.chm11662455
f1_keywords:
- vis_sdr.chm11662455
ms.prod: visio
api_name:
- Visio.Window.SelectionForDragCopy
ms.assetid: e34de916-5dc4-b9af-70b3-7c68340e2afb
ms.date: 06/08/2017
---


# Window.SelectionForDragCopy Property (Visio)

Returns the  **[Selection](selection-object-visio.md)** object that represents the collection of shapes that will participate in drag or copy operations, based on the current selection. Read-only.


## Syntax

 _expression_ . **SelectionForDragCopy**

 _expression_ A variable that represents a **[Window](window-object-visio.md)** object.


### Return Value

 **Selection**


## Remarks

The  **Selection** object that **SelectionForDragCopy** returns includes any unselected members of selected containers and lists, and unselected callouts associated with selected target shapes; all of these will also participate in the drag or copy operation.


