---
title: InvisibleApp.IsUndoingOrRedoing Property (Visio)
keywords: vis_sdr.chm17513760
f1_keywords:
- vis_sdr.chm17513760
ms.prod: visio
api_name:
- Visio.InvisibleApp.IsUndoingOrRedoing
ms.assetid: 2c13c2d7-0c22-c9f2-44e6-bb442acb2ca4
ms.date: 06/08/2017
---


# InvisibleApp.IsUndoingOrRedoing Property (Visio)

Determines whether the current event handler is being called as a result of an  **Undo** or **Redo** action in the application. Read-only.


## Syntax

 _expression_ . **IsUndoingOrRedoing**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Boolean


## Remarks

The  **IsUndoingOrRedoing** property returns **True** when the application is firing events related to an **Undo** or **Redo** action that the user has initiated in the user interface, or that an Automation client has initiated by calling the **Undo** or **Redo** method of an **Application** object.

When the application calls an event handler, the event has a "cause." If that cause is a user action or another event handler, it is legitimate to perform undoable actions during the course of handling that event. However, if the cause of the event firing is an  **Undo** or **Redo** action, the event handler should not perform undoable actions. Doing so eliminates the ability to redo an action.

You will typically only perform undoable actions inside an event handler when this property is  **False** . You can perform undoable actions when the flag is **True** , but the redo queue is destroyed.


## Example

This example adds a shape to a drawing and then checks whether the shape has been added as a result of an  **Undo** or **Redo** action.


```vb
 
Public Sub IsUndoingOrRedoing_Example() 
 
Dim vsoShape As Visio.Shape 
 
 'Draw a shape to trigger the ShapeAdded event 
 Set vsoShape = ActivePage.DrawRectangle(1,2,2,1) 
 
End Sub
```


```vb
 
Sub Document_ShapeAdded(ByVal Shape As IVShape) 
 
 If Not Application.IsUndoingOrRedoing Then 
 
 Debug.Print "Shape was added by new action." 
 
 End If 
 
End Sub
```


