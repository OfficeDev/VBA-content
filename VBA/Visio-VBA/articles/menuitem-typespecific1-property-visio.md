---
title: MenuItem.TypeSpecific1 Property (Visio)
keywords: vis_sdr.chm12914600
f1_keywords:
- vis_sdr.chm12914600
ms.prod: visio
api_name:
- Visio.MenuItem.TypeSpecific1
ms.assetid: fa0218de-5644-f2f0-9cad-d4d927349e00
ms.date: 06/08/2017
---


# MenuItem.TypeSpecific1 Property (Visio)

Gets or sets the type of a menu item. Read/write.


## Syntax

 _expression_ . **TypeSpecific1**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The value of an object's  **TypeSpecific1** property depends on the value of its **CntrlType** property.



|** CntrlType value**|** TypeSpecific1 value**|
|:-----|:-----|
| **visCtrlTypeBUTTON**|Any constant prefixed with  **visIconIX** that is declared by the Visio type library.|
| **visCtrlTypeCOMBOBOX**|Zero (0).|
| **visCtrlTypeEDITBOX**|Zero (0).|
| **visCtrlTypeLABEL**|The  **TypeSpecific1** property is not used.|

