---
title: Menu.ActionText Property (Visio)
keywords: vis_sdr.chm13113015
f1_keywords:
- vis_sdr.chm13113015
ms.prod: visio
api_name:
- Visio.Menu.ActionText
ms.assetid: 27d58281-5c00-56dd-25a4-2f65965daac2
ms.date: 06/08/2017
---


# Menu.ActionText Property (Visio)

Gets or sets the action text for a menu. Read/write.


## Syntax

 _expression_ . **ActionText**

 _expression_ A variable that represents a **Menu** object.


### Return Value

String


## Remarks

Action text is a string that describes the action on the  **Undo**,  **Redo**, and  **Repeat** menu items on the **Quick Access** toolbar.

If the  **ActionText** property is empty and the object's **CmdNum** property is set to one of the Microsoft Visio built-in command IDs, the item uses the default action text from the built-in Visio user interface.


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.


