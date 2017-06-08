---
title: MenuSet.Position Property (Visio)
keywords: vis_sdr.chm13314095
f1_keywords:
- vis_sdr.chm13314095
ms.prod: visio
api_name:
- Visio.MenuSet.Position
ms.assetid: 2e970661-b8d6-a886-ad26-89759272af9d
ms.date: 06/08/2017
---


# MenuSet.Position Property (Visio)

Gets or sets the position of an object. Read/write.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **MenuSet** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Constants that represent possible  **Position** property values are listed below. They are also declared by the Visio type library in **VisUIBarPosition** .



|** Constant**|** Value**|
|:-----|:-----|
| **visBarLeft**|0|
| **visBarTop**|1|
| **visBarRight**|2|
| **visBarBottom**|3|
| **visBarFloating**|4|
| **visBarPopup**|5|
| **visBarMenu**|6|

