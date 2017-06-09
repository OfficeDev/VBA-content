---
title: UIObject.UpdateUI Method (Visio)
keywords: vis_sdr.chm14916640
f1_keywords:
- vis_sdr.chm14916640
ms.prod: visio
api_name:
- Visio.UIObject.UpdateUI
ms.assetid: d5aefb7a-5d6f-5835-4c38-521aeceea289
ms.date: 06/08/2017
---


# UIObject.UpdateUI Method (Visio)

Causes Microsoft Visio to display changes to the user interface represented by a  **UIObject** object.


## Syntax

 _expression_ . **UpdateUI**

 _expression_ A variable that represents a **UIObject** object.


### Return Value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The  **UpdateUI** method updates the Visio user interface with changes made to a **UIObject** object during a session. Use the **CustomMenus** or **CustomToolbars** property of an **Application** object or **Document** object to obtain the **UIObject** object.


