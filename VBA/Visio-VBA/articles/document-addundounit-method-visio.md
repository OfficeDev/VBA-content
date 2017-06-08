---
title: Document.AddUndoUnit Method (Visio)
keywords: vis_sdr.chm10516075
f1_keywords:
- vis_sdr.chm10516075
ms.prod: visio
api_name:
- Visio.Document.AddUndoUnit
ms.assetid: 3b9d903a-8854-fa64-a9c5-85ac71d58f54
ms.date: 06/08/2017
---


# Document.AddUndoUnit Method (Visio)

Adds an object that supports the  **IOleUndoUnit** or **IVBUndoUnit** interface to the Microsoft Visio undo queue.


## Syntax

 _expression_ . **AddUndoUnit**( **_pUndoUnit_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pUndoUnit_|Required| **[UNKNOWN]**|A reference to an object that supports the  **IOleUndoUnit** or **IVBUndoUnit** interface.|

### Return Value

Nothing


## Remarks

For information about implementing the  **IOleUndoUnit** interface on your object, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network. For information about implementing the **IVBUndoUnit** interface, see Developing Microsoft Visio Solutions on MSDN.


## Example

The following procedure shows how to use the  **AddUndoUnit** method to add an object to the Visio undo queue. When a shape is added to the active document, the procedure checks to see if it was added as a result of an undo or redo action, and if not, it adds an Undo unit.

This procedure is a member of class  **clsParticipateInUndo** , which is defined in one of two related class modules in the Code Samples Library in the Visio SDK, and is not intended to be run independently. (The other class module defines class **clsVBUndoUnits** .) For more information on these class modules, see the Visio SDK on MSDN.




```vb
 
Private Sub mvsoDocument_ShapeAdded(ByVal vsoShape As IVShape) 
 
 Dim VBUndoUnit As clsVBUndoUnits 
 
 On Error GoTo mvsoDocument_ShapeAdded_Err 
 
 If Not (mvsoApplication Is Nothing) Then 
 
 If Not msvoApplication.IsUndoingOrRedoing Then 
 
 'Increment the count of undoable actions. 
 IncrementModuleVar 
 Debug.Print "Original Do: GetModuleVar = " &; GetModuleVar 
 
 'Instantiate clsVBUndoUnit, a 
 'class that implements Visio.IVBUndoUnit. 
 Set VBUndoUnit = New clsVBUndoUnits 
 
 'Pass the current instance of the class 
 'of which this procedure is a member, 
 'clsParticipateInUndo, to the Undo unit. 
 VBUndoUnit.SetModelObject Me 
 
 'Add an Undo unit. 
 mvsoApplication.AddUndoUnit VBUndoUnit 
 
 End If 
 
 End If 
 
Exit Sub 
 
mvsoDocument_ShapeAdded_Err: 
 
 MsgBox Err.Description 
 
End Sub
```


