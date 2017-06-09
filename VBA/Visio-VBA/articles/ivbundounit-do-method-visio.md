---
title: IVBUndoUnit.Do Method (Visio)
keywords: vis_sdr.chm17360155
f1_keywords:
- vis_sdr.chm17360155
ms.prod: visio
api_name:
- Visio.IVBUndoUnit.Do
ms.assetid: 3d33e1fe-328a-0337-412a-861b3e19d8b2
ms.date: 06/08/2017
---


# IVBUndoUnit.Do Method (Visio)

Called by the Undo Manager to tell an undo unit to perform its action.


## Syntax

 _expression_ . **Do**( **_pMgr_** )

 _expression_ A variable that represents an **IVBUndoUnit** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pMgr_|Required| **[IVBUNDOMANAGER]**|A pointer to an  **IVBUndoManager** interface.|

### Return Value

Nothing


## Remarks

If you are creating an undo unit for your solution, the  **Do** method is one of the procedures that you must implement. It provides the actions that are required to undo and redo actions.

 If you are creating a single object for both undoing and redoing, the **Do** method maintains the undo/redo state and adds an undo unit to the opposite stack.

If the  **Do** method is passed a **Nothing** pointer, the unit should carry out the undo action but should not place anything on the undo or redo stack.

For more information about the  **Do** method and using the **IVBUndoUnit** interface to create undo units, search for "creating undo units" on MSDN.


