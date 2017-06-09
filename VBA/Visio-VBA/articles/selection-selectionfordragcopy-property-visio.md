---
title: Selection.SelectionForDragCopy Property (Visio)
keywords: vis_sdr.chm11162455
f1_keywords:
- vis_sdr.chm11162455
ms.prod: visio
api_name:
- Visio.Selection.SelectionForDragCopy
ms.assetid: f7e6e87a-c904-6008-fdde-4d5cb124351c
ms.date: 06/08/2017
---


# Selection.SelectionForDragCopy Property (Visio)

Returns the  **[Selection](selection-object-visio.md)** object that represents the collection of shapes that will participate in drag or copy operations, based on the current selection. Read-only.


## Syntax

 _expression_ . **SelectionForDragCopy**

 _expression_ A variable that represents a **Selection** object.


### Return Value

 **Selection**


## Remarks

The  **Selection** object that **SelectionForDragCopy** returns includes any unselected members of selected containers and lists, and unselected callouts that are associated with selected target shapes; all of these will also participate in the drag or copy operation.


