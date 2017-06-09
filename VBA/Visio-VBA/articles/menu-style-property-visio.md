---
title: Menu.Style Property (Visio)
keywords: vis_sdr.chm13151150
f1_keywords:
- vis_sdr.chm13151150
ms.prod: visio
api_name:
- Visio.Menu.Style
ms.assetid: 25d36a5a-d109-bd60-7fea-6f22eba8b5bb
ms.date: 06/08/2017
---


# Menu.Style Property (Visio)

Determines whether a menu shows an icon, a caption, or some combination. Read/write.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Menu** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Possible values for the  **Style** property are listed in the following table. These constants are declared by the Visio type library in **VisUIButtonStyle** .



|** Constant**|** Value**|
|:-----|:-----|
| **visButtonAutomatic**|0|
| **visButtonCaption**|1|
| **visButtonIcon**|2|
| **visButtonIconandCaption**|3|

