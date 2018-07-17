---
title: MenuSet.SetID Property (Visio)
keywords: vis_sdr.chm13314315
f1_keywords:
- vis_sdr.chm13314315
ms.prod: visio
api_name:
- Visio.MenuSet.SetID
ms.assetid: d1971b55-a03d-dbd7-608f-8b7c88f526c6
ms.date: 06/08/2017
---


# MenuSet.SetID Property (Visio)

Returns the set ID of a  **MenuSet** object in its collection. Read-only.


## Syntax

 _expression_ . **SetID**

 _expression_ A variable that represents a **MenuSet** object.


### Return Value

Long


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Each  **MenuSet** object has a set ID that corresponds to a Microsoft Visio window context. For **MenuSet** objects, set IDs also correspond to shortcut menu sets.

You can retrieve an object from its collection by passing the object's set ID to the  **ItemAtID** property. You can also set the set ID of an object by using the **AddAtID** method.

Valid set ID values are declared by the Visio type library in  **[VisUIObjSets](visuiobjsets-enumeration-visio.md)** .


