---
title: ToolbarSet.SetID Property (Visio)
keywords: vis_sdr.chm13914315
f1_keywords:
- vis_sdr.chm13914315
ms.prod: visio
api_name:
- Visio.ToolbarSet.SetID
ms.assetid: db1f1cf5-f9eb-a118-132d-9ac878db6632
ms.date: 06/08/2017
---


# ToolbarSet.SetID Property (Visio)

Returns the set ID of an  **ToolbarSet** object in its collection. Read-only.


## Syntax

 _expression_ . **SetID**

 _expression_ A variable that represents a **ToolbarSet** object.


### Return Value

Long


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Each  **ToolbarSet** object has a set ID that corresponds to a Microsoft Visio window context. For **ToolbarSet** objects, they also correspond to drop-down menus under toolbar buttons (such as **Fill Color** or **Line Weight**).

You can retrieve an object from its collection by passing the object's set ID to the  **ItemAtID** property. You can also set the set ID of an object by using the **AddAtID** method.

Valid set ID values are declared by the Visio type library in  **[VisUIObjSets](visuiobjsets-enumeration-visio.md)** .


