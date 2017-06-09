---
title: IVBUndoUnit.OnNextAdd Method (Visio)
keywords: vis_sdr.chm17360160
f1_keywords:
- vis_sdr.chm17360160
ms.prod: visio
api_name:
- Visio.IVBUndoUnit.OnNextAdd
ms.assetid: a5504398-75a9-06be-346c-3afd85ce708e
ms.date: 06/08/2017
---


# IVBUndoUnit.OnNextAdd Method (Visio)

Notifies an undo unit that another undo unit has been added to the undo stack. Returns  **Nothing** .


## Syntax

 _expression_ . **OnNextAdd**

 _expression_ A variable that represents an **IVBUndoUnit** object.


### Return Value

Nothing


## Remarks

If you are creating an undo unit for your solution, the  **OnNextAdd** method is one of the procedures that you must implement. It gets called when the next undo unit in the same undo scope gets added to the undo stack.

When an undo unit receives an  **OnNextAdd** notification, it communicates back to the creating object that it can no longer insert data into this undo unit.

For more information about the using the  **OnNextAdd** method and using the **IVBUndoUnit** interface to create undo units, search for "creating undo units" on MSDN.


