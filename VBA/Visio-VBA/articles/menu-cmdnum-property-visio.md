---
title: Menu.CmdNum Property (Visio)
keywords: vis_sdr.chm13113255
f1_keywords:
- vis_sdr.chm13113255
ms.prod: visio
api_name:
- Visio.Menu.CmdNum
ms.assetid: 13754873-94bd-3497-829c-374aec3615da
ms.date: 06/08/2017
---


# Menu.CmdNum Property (Visio)

Gets or sets the command ID associated with a menu. Read/write.


## Syntax

 _expression_ . **CmdNum**

 _expression_ A variable that represents a **Menu** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Valid command IDs are declared by the Visio type library in  **[VisUICmds](visuicmds-enumeration-visio.md)** . They have the prefix **visCmd** .


