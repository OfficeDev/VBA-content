---
title: ToolbarItem Object (Visio)
keywords: vis_sdr.chm10275
f1_keywords:
- vis_sdr.chm10275
ms.prod: visio
api_name:
- Visio.ToolbarItem
ms.assetid: 2f0798cf-f31e-e213-d9db-325d58a77e96
ms.date: 06/08/2017
---


# ToolbarItem Object (Visio)

Represents one item in a  **Toolbar** object. A **ToolbarItem** object can represent a button, combo box, or any other item on the Microsoft Visio toolbars.


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The index of the  **ToolbarItem** object within the **ToolbarItems** collection corresponds to its position on the toolbar, starting with zero (0) for the item farthest to the left if the toolbars are arranged horizontally.

Beginning with Microsoft Visio 2002, use the  **BeginGroup** property to create spaces on a toolbar.


