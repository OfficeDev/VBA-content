---
title: DrawingControl Object (Visio)
keywords: vis_sdr.chm0
f1_keywords:
- vis_sdr.chm0
ms.prod: visio
api_name:
- Visio.DrawingControl
ms.assetid: ad7c6abf-5bbd-5b84-4a63-eceaf90991a8
ms.date: 06/08/2017
---


# DrawingControl Object (Visio)

A programmable ActiveX control that enables you to build Microsoft Visio functionality into programs you create in Microsoft Visual Studio and other development platforms.


## Remarks

Use the  **Document** property of the **DrawingControl** object to get the Visio **Document** object associated with the instance of the Microsoft Visio Drawing Control and thereby gain access to the Visio object model.

Use the  **HostID** property of the **DrawingControl** object to assign a GUID or other string representation of the container application to a registry key.

Use the  **NegotiateMenus** and **Negotiate Toolbars** properties of the **DrawingControl** object to determine whether Visio menus and toolbars are merged with those of the host container application in the Visio Drawing Control, and to enable programmatic customizing of Visio menus and toolbars.


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Use the  **PageSizingBehavior** property of the **DrawingControl** object to specify how the behavior of the control changes as the control is resized, with respect to the drawing page and any shapes on it.

Use the  **Src** property of the **DrawingControl** object to specify the Visio drawing to appear in the Visio Drawing Control.

Use the  **Window** property of the **DrawingControl** object to get the Visio **Window** object associated with the instance of the Visio Drawing Control and thereby gain access to the Visio object model.

The  **DrawingControl** object has no default property.


