---
title: ToolbarItem.ActionText Property (Visio)
keywords: vis_sdr.chm13513015
f1_keywords:
- vis_sdr.chm13513015
ms.prod: visio
api_name:
- Visio.ToolbarItem.ActionText
ms.assetid: 0bb1f3a9-013d-4164-6de3-557dcf64ca92
ms.date: 06/08/2017
---


# ToolbarItem.ActionText Property (Visio)

Gets or sets the action text for a toolbar item. Read/write. 


## Syntax

 _expression_ . **ActionText**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Action text is a string that describes the action on the  **Undo**,  **Redo**, and  **Repeat** menu items on the **Edit** menu.

If the  **ActionText** property is empty and the object's **CmdNum** property is set to one of the Microsoft Visio built-in command IDs, the item uses the default action text from the built-in Visio user interface.


