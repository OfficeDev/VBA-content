---
title: Toolbar.RowIndex Property (Visio)
keywords: vis_sdr.chm13714255
f1_keywords:
- vis_sdr.chm13714255
ms.prod: visio
api_name:
- Visio.Toolbar.RowIndex
ms.assetid: f55616ce-73a0-332b-ece0-f9f83ef43547
ms.date: 06/08/2017
---


# Toolbar.RowIndex Property (Visio)

Gets or sets the docking order of a  **Toolbar** object in relation to other items in the same docking area. Read/write.


## Syntax

 _expression_ . **RowIndex**

 _expression_ A variable that represents a **Toolbar** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Objects that have lower numbers are docked first. Several items can share the same row index. If two or more items share the same row index, the item most recently assigned is displayed first in its group.

Constants that represent the first and last positions (see the following table) are declared by the Visio type library in member  **[VisUIBarRow](visuibarrow-enumeration-visio.md)** .



|** Constant**|** Value**|
|:-----|:-----|
| **visBarRowFirst**|0|
| **visBarRowLast**|-1|

