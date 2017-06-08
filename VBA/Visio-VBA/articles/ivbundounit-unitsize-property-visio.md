---
title: IVBUndoUnit.UnitSize Property (Visio)
keywords: vis_sdr.chm17360165
f1_keywords:
- vis_sdr.chm17360165
ms.prod: visio
api_name:
- Visio.IVBUndoUnit.UnitSize
ms.assetid: 4e6fac31-60d2-e6d5-324d-c593b0456c95
ms.date: 06/08/2017
---


# IVBUndoUnit.UnitSize Property (Visio)

Returns the size of the undo unit in memory, in bytes. Read-only.


## Syntax

 _expression_ . **UnitSize**

 _expression_ A variable that represents a **IVBUndoUnit** object.


### Return Value

Long


## Remarks

If you are creating an undo unit for your solution, the  **UnitSize** property is one of the members of **IVBUndoUnit** that you must implement. The Visio engine may use the memory size returned by the **UnitSize** property to decide whether to clear the undo queue.

 The **UnitSize** value is optional, and you can set the value to zero (0) when you implement the property.

For more information about the  **UnitSize** property and using the **IVBUndoUnit** interface to create undo units, search for "Creating Undo Units" on MSDN, the Microsoft Developer Network.


