---
title: Menu.FaceID Property (Visio)
keywords: vis_sdr.chm13113495
f1_keywords:
- vis_sdr.chm13113495
ms.prod: visio
api_name:
- Visio.Menu.FaceID
ms.assetid: 03270afe-84ea-d21d-9077-5967dfce3550
ms.date: 06/08/2017
---


# Menu.FaceID Property (Visio)

Gets or sets the icon for an item. Read/write.


## Syntax

 _expression_ . **FaceID**

 _expression_ A variable that represents a **Menu** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You can use any of the constants prefixed with  **visIconIX** that are declared by the Visio type library in **[VisUIIconIDs](visuiiconids-enumeration-visio.md)** .

The  **FaceID** property is the same as the **TypeSpecific1** property when the **CtrlType** property is type **visCtrlTypeBUTTON** , which is declared in the Visio type library in **[VisUICtrlTypes](visuictrltypes-enumeration-visio.md)** .


