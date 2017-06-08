---
title: IVBUndoUnit Object (Visio)
keywords: vis_sdr.chm60145
f1_keywords:
- vis_sdr.chm60145
ms.prod: visio
ms.assetid: 397d8ea4-50ec-970a-61bb-ca61b2ae84e3
ms.date: 06/08/2017
---


# IVBUndoUnit Object (Visio)

The interface on an undo unit in Microsoft Visio. An undo unit encapsulates the information necessary to undo or redo a single action.


## Remarks

The default property of  **IVBUndoUnit** is **Description** .

You can use the  **IVBUndoUnit** interface in Microsoft Visual Basic for Applications (VBA) or Microsoft Visual Basic to create your own undo units for the Visio undo manager. To create an undo unit, you must implement this interface, along with all of its public procedures, in a class module that you insert into your project.

For more information about using the  **IVBUndoUnit** interface to create undo units, search for "creating undo units" on MSDN, the Microsoft Developer Network.


### IVBUndoUnit methods and properties in VTable order



|** IUnknown Methods**|** Description**|
|:-----|:-----|
| **QueryInterface**| Returns a pointer to a specified interface.|
| **AddRef**| Increments the reference count.|
| **Release**| Decrements the reference count.|
| IVBUndoUnit Methods| Description|
| **Do**| Instructs the undo unit to carry out its action.|
| **OnNextAdd**| Notifies the last undo unit in the collection that a new unit has been added.|
| IVBUndoUnit Properties| Description|
| **Description**| Read-only. Describes the undo action.|
| **UnitTypeCLSID**| Read-only. Returns the CLSID and a type identifier for the undo unit.|
| **UnitTypeLong**| Read-only. Returns a **Long** that can be used to identify the undo unit.|
| **UnitSize**| Size in bytes. Used to measure how much memory undo information is using.|

