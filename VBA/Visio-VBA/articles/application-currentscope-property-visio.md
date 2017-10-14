---
title: Application.CurrentScope Property (Visio)
keywords: vis_sdr.chm10013340
f1_keywords:
- vis_sdr.chm10013340
ms.prod: visio
api_name:
- Visio.Application.CurrentScope
ms.assetid: a45fd841-efb4-90b6-65fb-21f9f8e8ea0c
ms.date: 06/08/2017
---


# Application.CurrentScope Property (Visio)

Determines the ID of the scope that causes an event to fire. Read-only.


## Syntax

 _expression_ . **CurrentScope**

 _expression_ A variable that represents an **Application** object.


### Return Value

Long


## Remarks

Returns  **visScopeIDInvalid** (-1) if a scope isn't open. The scope ID could be an internal Microsoft Visio scope ID that corresponds to a Visio command or an external scope ID passed to an Automation client by the **BeginUndoScope** method.

The recipients of an event consider a scope open if the  **EnterScope** event has fired but the **ExitScope** event has not fired.

To determine if the event queue firing is related to a particular scope internal to the application or one opened and closed by an Automation client, use the  **IsInScope** property.


## Example

This example shows how to use the  **CurrentScope** property to determine the ID of the current scope.


```vb
Private WithEvents vsoApplication As Visio.Application 
Private lngScopeID As Long 
 
Public Sub ScopeActions() 
 
 Dim vsoShape As Visio.Shape 
 
 'Set the module level application variable to 
 'trap Application level events. 
 Set vsoApplication = Application 
 
 'Begin a scope, set the module level variable. 
 lngScopeID = Application.BeginUndoScope("Draw Shapes") 
 
 'Draw three shapes. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 ActivePage.DrawOval 3, 4, 4, 3 
 ActivePage.DrawLine 4, 5, 5, 4 
 
 'Change a cell (which would trigger a cell changed event). 
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
 Debug.Print "Entering current scope " &; nScopeID 
 Else 
 Debug.Print "Enter Scope " &; bstrDescription &; "(" &; nScopeID &; ")" 
 End If 
 
End Sub 
 
Private Sub vsoApplication_ExitScope(ByVal app As IVApplication, _ 
 ByVal nScopeID As Long, _ 
 ByVal strDescription As String, _ 
 ByVal bErrOrCancelled As Boolean) 
 
 If vsoApplication.CurrentScope = lngScopeID Then 
 Debug.Print "Exiting current scope " &; nScopeID 
 Else 
 Debug.Print "ExitScope " &; bstrDescription &; "(" &; nScopeID &; ")" 
 End If 
 
End Sub
```


