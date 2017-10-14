---
title: Menus.Parent Property (Visio)
keywords: vis_sdr.chm13214040
f1_keywords:
- vis_sdr.chm13214040
ms.prod: visio
api_name:
- Visio.Menus.Parent
ms.assetid: 1dffd96d-d53c-874d-405b-c8f9de9ae459
ms.date: 06/08/2017
---


# Menus.Parent Property (Visio)

Determines the parent of an object. Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **Menus** object.


### Return Value

MenuSet


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

In general, an object's parent is the object that contains it. For example, the parent of a  **Menu** object is the **Menus** collection that contains the **Menu** object.


