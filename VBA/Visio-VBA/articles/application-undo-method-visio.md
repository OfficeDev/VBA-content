---
title: Application.Undo Method (Visio)
keywords: vis_sdr.chm10016620
f1_keywords:
- vis_sdr.chm10016620
ms.prod: visio
api_name:
- Visio.Application.Undo
ms.assetid: 728d9af0-c9f2-c3ff-5ed3-a20e8a507a6a
ms.date: 06/08/2017
---


# Application.Undo Method (Visio)

Reverses the most recent undo unit, if the undo unit can be reversed.


## Syntax

 _expression_ . **Undo**

 _expression_ A variable that represents an **Application** object.


### Return Value

Nothing


## Remarks

Use the  **Undo** method to reverse actions one undo unit at a time.

The number of times that code can call the  **Undo** method depends on whether or not the code is executing in the scope of an open undo unit. Code runs in the scope of an open undo unit if it is:




- A macro or add-on invoked by the Microsoft Visio user interface.
    
- In an event handler responding to a Visio event other than the  **VisioIsIdle** event.
    
- In a user-created undo scope.
    


If code is not executing in the scope of an open undo unit, it can call the  **Undo** method for each undo unit presently on the Visio undo stack. You can set the maximum number of units on the undo stack (20 is the default) on the **Advanced** tab of the **Visio Options** dialog box (click the **File** tab, and then click **Options**). If the number of calls to the  **Undo** method exceeds the number of undo units on the stack, no action is taken and the **Undo** method raises no exception.

If code is executing in the scope of an open undo unit, it can call the  **Undo** method once for each operation in the open undo unit. If there are additional calls to the **Undo** method, it raises an exception and takes no action. For example, if code in a macro performs two operations, it can call the **Undo** method twice. If the macro calls the **Undo** method a third time, the **Undo** method raises an exception.

Code that calls the  **Undo** method from within the scope of an undo unit cannot call the **Redo** method to reverse the action. The **Redo** method can only be called when there are no open undo units.

The  **Undo** method also raises an exception if the Visio instance is presently performing an undo or redo. To determine whether the Visio instance is undoing or redoing, use the **IsUndoingOrRedoing** property.

You can call the  **Undo** method from the **VisioIsIdle** event handler because the **VisioIsIdle** event can only fire when the **IsUndoingOrRedoing** property is **False** . You can also call the **Undo** method from code not invoked by the Visio instance, for example, code invoked from the Visual Basic Editor or from an external program.

You can undo most actions, but not all. Use the  **Redo** method to reverse the effect of the **Undo** method.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to undo and redo actions.


```vb
 
Public Sub Undo_Example()  
 
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


