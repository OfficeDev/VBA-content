---
title: MenuItem.Index Property (Visio)
keywords: vis_sdr.chm12913695
f1_keywords:
- vis_sdr.chm12913695
ms.prod: visio
api_name:
- Visio.MenuItem.Index
ms.assetid: 354a86ac-b96f-1d3a-3d6e-3cd319b22f16
ms.date: 06/08/2017
---


# MenuItem.Index Property (Visio)

Gets the ordinal position of a  **MenuItem** object in the **MenuItems** collection. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

Long


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Most collections are indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so forth. The index of the last element in a collection is the same as the value of that collection's  **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.

There are some exceptions. The  **Colors** collection is indexed starting with 0.

These collections are also indexed starting with 0:  **AccelItems** , **AccelTables** , **MenuSets** , **MenuItems** , **Menus** , **ToolbarItems** , **Toolbars** , and **ToolbarSets** .


