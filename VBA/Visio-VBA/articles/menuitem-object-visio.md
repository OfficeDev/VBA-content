---
title: MenuItem Object (Visio)
keywords: vis_sdr.chm10155
f1_keywords:
- vis_sdr.chm10155
ms.prod: visio
api_name:
- Visio.MenuItem
ms.assetid: 7161bf25-fde8-09d6-0c10-52a65f86feba
ms.date: 06/08/2017
---


# MenuItem Object (Visio)

Represents a single menu item on a Microsoft Visio menu.


## Version Information

Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.


## Remarks


 **Note**  

The default property of  **MenuItem** is **Caption** .

A  **MenuItem** object contains all the information it needs to display the menu item and launch the appropriate Visio command or add-on. It also contains text for the **Undo**,  **Redo**, and  **Repeat** menu items and error messages.

The index of a  **MenuItem** object within the **MenuItems** collection corresponds to the menu item's position from top to bottom on the menu or submenu, starting with zero (0).

If the menu item displays a submenu, the  **MenuItem** object has a **MenuItems** collection that represents items on the submenu. The **MenuItem** object's **Caption** property contains the submenu title. Most of the other properties of the **MenuItem** object are ignored, because this object serves much the same role as a **Menu** object.


