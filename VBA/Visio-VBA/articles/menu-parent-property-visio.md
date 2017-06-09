---
title: Menu.Parent Property (Visio)
keywords: vis_sdr.chm13114040
f1_keywords:
- vis_sdr.chm13114040
ms.prod: visio
api_name:
- Visio.Menu.Parent
ms.assetid: d402cd11-c59b-b4aa-883f-24e9c096f548
ms.date: 06/08/2017
---


# Menu.Parent Property (Visio)

Determines the parent of an object. Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **Menu** object.


### Return Value

Menus


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

In general, an object's parent is the object that contains it. For example, the parent of a  **Menu** object is the **Menus** collection that contains the **Menu** object.


