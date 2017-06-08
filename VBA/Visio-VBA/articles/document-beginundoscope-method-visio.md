---
title: Document.BeginUndoScope Method (Visio)
keywords: vis_sdr.chm10516085
f1_keywords:
- vis_sdr.chm10516085
ms.prod: visio
api_name:
- Visio.Document.BeginUndoScope
ms.assetid: 4e0c99a3-3ac6-54f8-3e43-1c79224e09e1
ms.date: 06/08/2017
---


# Document.BeginUndoScope Method (Visio)

Starts a transaction with a unique scope ID for an instance of Microsoft Visio.


## Syntax

 _expression_ . **BeginUndoScope**( **_bstrUndoScopeName_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrUndoScopeName_|Required| **String**|The name of the scope; could appear in the Visio user interface.|

### Return Value

Long


## Remarks

If you need to know whether events you receive are the result of a particular operation that you initiated, use the  **BeginUndoScope** and **EndUndoScope** methods to wrap your operation. In your event handlers, use the **IsInScope** property to test whether the scope ID returned by the **BeginUndoScope** method is part of the current context. Make sure you clear the scope ID you stored from the **BeginUndoScope** property when you receive the **ExitScope** event with that ID.

You must balance calls to the  **BeginUndoScope** method with calls to the **EndUndoScope** method. If you call the **BeginUndoScope** method, you should call the **EndUndoScope** method as soon as you are finished with the actions that constitute your scope. Also, while actions to multiple documents should be robust within a single scope, closing a document may have the side effect of clearing the undo information for the currently open scope as well as clearing the undo and redo stacks. If that happens, passing _bCommit_ = **False** to **EndUndoScope** does not restore the undo information.

You can also use the  **BeginUndoScope** and **EndUndoScope** methods to add an action defined by an add-on to the Visio undo stream. This is useful when you are operating from modeless scenarios where the initiating agent is part of an add-on's user interface or a modeless programmatic action.


 **Note**  Most Visio actions are already wrapped in internal undo scopes, so add-ons running within the application do not need to call this method.


## Example

This example shows how to use the  **BeginUndoScope** method to start a transaction that has a unique scope ID for an instance of Visio.


```vb
 
Private WithEvents vsoApplication As Visio.Application 
Private lngScopeID As Long 
 
Public Sub BeginUndoScope_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 'Set the module-level application variable to 
 'trap application-level events. 
 Set vsoApplication = Application 
 
 'Begin a scope and set the module-level scope ID variable. 
 lngScopeID = Application.BeginUndoScope("Draw Shapes") 
 
 'Draw three shapes. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 ActivePage.DrawOval 3, 4, 4, 3 
 ActivePage.DrawLine 4, 5, 5, 4 
 
 'Change a cell to trigger the CellChanged event. 
 vsoShape.Cells("Width").Formula = 5 
 
 'End and commit this scope. 
 Application.EndUndoScope lngScopeID, True 
 
 End Sub 
 
 Private Sub vsoApplication_CellChanged(ByVal Cell As IVCell) 
 
 'Check to see if this cell change is the result of something 
 'happening within the scope. 
 If vsoApplication.IsInScope(lngScopeID) Then 
 Debug.Print Cell.Name &; " changed in scope "; lngScopeID 
 End If 
 
End Sub 
 
Private Sub vsoApplication_EnterScope(ByVal app As IVApplication, _ 
 ByVal nScopeID As Long, _ 
 ByVal bstrDescription As String) 
 
 If vsoApplication.CurrentScope = lngScopeID Then 
 Debug.Print "Entering my scope " &; nScopeID 
 Else 
 Debug.Print "Enter Scope " &; bstrDescription &; "(" &; nScopeID &; ")" 
 End If 
 
End Sub 
 
Private Sub vsoApplication_ExitScope(ByVal app As IVApplication, _ 
 ByVal nScopeID As Long, _ 
 ByVal bstrDescription As String, _ 
 ByVal bErrOrCancelled As Boolean) 
 
 If vsoApplication.CurrentScope = lngScopeID Then 
 Debug.Print "Exiting my scope " &; nScopeID 
 Else 
 Debug.Print "Exit Scope " &; bstrDescription &; "(" &; nScopeID &; ")" 
 End If 
 
End Sub
```


