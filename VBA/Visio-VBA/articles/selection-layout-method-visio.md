---
title: Selection.Layout Method (Visio)
keywords: vis_sdr.chm11116385
f1_keywords:
- vis_sdr.chm11116385
ms.prod: visio
api_name:
- Visio.Selection.Layout
ms.assetid: 58ff8c1f-92b3-2473-d786-28e64e7c5586
ms.date: 06/08/2017
---


# Selection.Layout Method (Visio)

Lays out the shapes and/or reroutes the connectors for the page, master, group, or selection.


## Syntax

 _expression_ . **Layout**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Using the  **Layout** method is equivalent to clicking **Re-Layout Page** in the **Layout** group on the **Design** tab and then clicking **More Layout Options** and configuring options in the **Configure Layout** dialog box.

Behavior of the  **Layout** method can be influenced by setting the formulas or results of cells in the Page Layout and Shape Layout ShapeSheet sections of the page, master, or group to be laid out. You can infer how these cells influence the behavior of the **Layout** method by examining the effect of various **Configure Layout** dialog box options on the values of these cells.

To lay out a subset of the shapes of a page, master, or group, establish a  **Selection** object in which the shapes to be laid out are selected, and then call the **Layout** method. If the **Layout** method is performed on a **Selection** object and the object has no shapes selected, all shapes in the page, master, or group of the selection are laid out.


