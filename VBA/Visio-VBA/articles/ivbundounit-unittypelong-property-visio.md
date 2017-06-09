---
title: IVBUndoUnit.UnitTypeLong Property (Visio)
keywords: vis_sdr.chm17360175
f1_keywords:
- vis_sdr.chm17360175
ms.prod: visio
api_name:
- Visio.IVBUndoUnit.UnitTypeLong
ms.assetid: 4fb63748-baf1-3360-f143-52de4c24c16d
ms.date: 06/08/2017
---


# IVBUndoUnit.UnitTypeLong Property (Visio)

Identifies an undo unit by a  **Long** . Read-only.


## Syntax

 _expression_ . **UnitTypeLong**

 _expression_ A variable that represents a **IVBUndoUnit** object.


### Return Value

Long


## Remarks

If you are creating an undo unit for your solution, the  **UnitTypeLong** property is one of the members of **IVBUndoUnit** that you must implement. You can use the **UnitTypeLong** value to identify your undo units.

 The **UnitTypeLong** value is optional, and you can and set the value to zero (0) when you implement the property.

For more information about the  **UnitTypeLong** property and using the **IVBUndoUnit** interface to create undo units, search for "creating undo units" on MSDN, the Microsoft Developer Network.


