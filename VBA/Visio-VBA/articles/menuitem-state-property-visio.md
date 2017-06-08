---
title: MenuItem.State Property (Visio)
keywords: vis_sdr.chm12914425
f1_keywords:
- vis_sdr.chm12914425
ms.prod: visio
api_name:
- Visio.MenuItem.State
ms.assetid: bcad69bc-790e-475d-ec4f-521a112393e3
ms.date: 06/08/2017
---


# MenuItem.State Property (Visio)

Determines a menu item's state, pressed or not pressed. Read/write.


## Syntax

 _expression_ . **State**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The  **State** property can be one of the following constants declared by the Visio type library in **VisUIButtonState** .



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visButtonUp**|0|Button is not pressed|
| **visButtonDown**|-1|Button is pressed|

