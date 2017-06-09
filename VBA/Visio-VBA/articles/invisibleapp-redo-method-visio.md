---
title: InvisibleApp.Redo Method (Visio)
keywords: vis_sdr.chm17516465
f1_keywords:
- vis_sdr.chm17516465
ms.prod: visio
api_name:
- Visio.InvisibleApp.Redo
ms.assetid: a72c916c-cf52-d9f3-b2ee-e0225f4b6940
ms.date: 06/08/2017
---


# InvisibleApp.Redo Method (Visio)

Reverses the most recent undo unit.


## Syntax

 _expression_ . **Redo**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

nothing


## Remarks

To reverse the effect of the  **Undo** method, use the **Redo** method. For example, if you clear an item and then use the **Undo** method to restore it, use the **Redo** method to clear the item again.

You cannot invoke the  **Redo** method from code that is executing inside the scope of an open undo unit. Code is in the scope of an open undo unit if it is one of the following:




- A macro or add-on invoked by the Microsoft Visio user interface.
    
- In an event handler responding to a Visio event other than the  **VisioIsIdle** event.
    
- In a user-created undo scope. If you call the  **Redo** method from code inside the scope of an open undo unit, it will raise an exception.
    


The  **Redo** method also raises an exception if the Visio instance is presently performing an undo or redo. To determine whether the Visio instance is undoing or redoing use the **IsUndoingOrRedoing** property.

You can call the  **Redo** method from the **VisioIsIdle** event handler because the **VisioIsIdle** event can only fire when the **IsUndoingOrRedoing** property is **False** . You can also call the **Redo** method from code not invoked by the Visio instance, for example, code invoked from the Visual Basic Editor or from an external program.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to undo and redo actions.


```vb
 
Public Sub Redo_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 'Draw a rectangle, use Undo to delete it, and 
 'then use Redo to redraw it. 
 Set vsoShape = ActivePage.DrawRectangle(1, 5, 5, 1) 
 
 'Delete the shape. 
 Visio.Application.Undo 
 
 'Bring it back. 
 Visio.Application.Redo 
 
End Sub
```


