---
title: ToolbarItem.BeginGroup Property (Visio)
keywords: vis_sdr.chm13551115
f1_keywords:
- vis_sdr.chm13551115
ms.prod: visio
api_name:
- Visio.ToolbarItem.BeginGroup
ms.assetid: fa1648ed-0876-d31d-6afe-b0278b0488f8
ms.date: 06/08/2017
---


# ToolbarItem.BeginGroup Property (Visio)

Determines whether a toolbar item appears at the beginning of a group of items on the toolbar. Read/write.


## Syntax

 _expression_ . **BeginGroup**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Boolean


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If you set the  **BeginGroup** property of a **ToolbarItem** object to **True** , a spacer is inserted into the toolbar preceding this item.


 **Note**  In Microsoft Visio 2000, the only way to create a separator in a menu or a spacer in a toolbar was to add a dummy item with a  **CmdNum** property of zero, a **Caption** property that contained "", and an empty **MenuItems** or **ToolbarItems** collection. This technique continues to work in later versions.


