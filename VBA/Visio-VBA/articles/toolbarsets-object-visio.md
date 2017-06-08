---
title: ToolbarSets Object (Visio)
keywords: vis_sdr.chm10295
f1_keywords:
- vis_sdr.chm10295
ms.prod: visio
api_name:
- Visio.ToolbarSets
ms.assetid: ddf79048-6585-81ab-b1c6-d7c4b0f0ff1b
ms.date: 06/08/2017
---


# ToolbarSets Object (Visio)

Includes a  **ToolbarSet** object for each window context that can display toolbars.


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

To retrieve a  **ToolbarSets** collection, use the **ToolbarSets** property of a **UIObject** object.

The default property of a  **ToolbarSets** collection is **Item** .

Unlike other Microsoft Visio collections, the  **ToolbarSets** collection is indexed starting with zero (0) rather than 1.

A  **ToolbarSet** object is identified in the **ToolbarSets** collection by its **SetID** property, which corresponds to a Visio window context. For a list of **SetID** values for **ToolbarSet** objects, see the **SetID** property.


