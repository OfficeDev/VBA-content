---
title: Selection.PrimaryItem Property (Visio)
keywords: vis_sdr.chm11114100
f1_keywords:
- vis_sdr.chm11114100
ms.prod: visio
api_name:
- Visio.Selection.PrimaryItem
ms.assetid: febdc4ec-d7db-7b4f-145b-aa9b23a2d5d2
ms.date: 06/08/2017
---


# Selection.PrimaryItem Property (Visio)

Returns the  **Shape** object that is a **Selection** object's primary item. Read-only.


## Syntax

 _expression_ . **PrimaryItem**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Shape


## Remarks

In a drawing window, the primary selected item is shown with green selection handles and non-primary selected items are shown with blue selection handles. The outcome of some operations is affected by which selected item is the primary item. For example, the  **Align Shapes** command aligns non-primary selected items with the primary selected item.

If a  **Selection** object contains no **Shape** objects, or the primary **Shape** object is one that isn't enumerated given the **Selection** object's **IterationMode** property, the **PrimaryItem** property returns **Nothing** and raises no exception.


