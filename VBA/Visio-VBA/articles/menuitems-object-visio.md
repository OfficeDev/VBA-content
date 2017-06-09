---
title: MenuItems Object (Visio)
keywords: vis_sdr.chm10160
f1_keywords:
- vis_sdr.chm10160
ms.prod: visio
api_name:
- Visio.MenuItems
ms.assetid: 7799eff9-5432-9c44-2e74-345479eef5b6
ms.date: 06/08/2017
---


# MenuItems Object (Visio)

 Contains a **MenuItem** object for each command on a Microsoft Visio menu.


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

To retrieve a  **MenuItems** collection, use the **MenuItems** property of a **Menu** object or a **MenuItem** object.

The default property of a  **MenuItems** collection is **Item** .

Unlike other Visio collections, the  **MenuItems** collection is indexed starting with zero (0) rather than 1.


