---
title: InvisibleApp.ExitScope Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ExitScope
ms.assetid: c035f0c2-af15-8557-6cac-0c3cd14d3599
ms.date: 06/08/2017
---


# InvisibleApp.ExitScope Event (Visio)

Queued when an internal command ends, or when an automation client exits a scope by using the  **EndUndoScope** method.


## Syntax

Private Sub  _expression_ _**ExitScope**( **_ByVal app As [IVAPPLICATION]_** , **_ByVal nScopeID As Long_** , **_ByVal bstrDescription As String_** , **_ByVal bErrOrCancelled As Boolean_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that contains the scope.|
| _nScopeID_|Required| **Long**|A language-independent number that describes the operation that just ended or the scope ID returned by the  **BeginUndoScope** method.|
| _bstrDescription_|Required| **String**| A textual description of the operation that changes in different language versions. Contains the UI description of a Visio operation or the description passed to the **BeginUndoScope** method.|
| _bErrOrCancelled_|Required| **Boolean**| **True** if there was an error during the scope or if the scope was canceled; **False** if there wasn't an error and it wasn't canceled.|

## Remarks

The  _nScopeID_ value returned in the case of a Visio operation is the equivalent of the command-related constants that begin with **visCmd*** .

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).

If you are handling this event from a program that receives a notification over a connection created using the  **AddAdvise method**, the  **ExitScope** event is one of a group of selected events that record extra information in the **EventInfo** property of the **Application** object.

The  **EventInfo** property returns _bstrDescription_, as described previously. In addition, the  _varMoreInfo_ argument to **VisEventProc** contains a string formatted as follows: [<nScopeID>;<bErrOrCancelled>;<bstrDescription>;<nHwndContext>], where _nHwndContext_ is the window handle (HWND) of the window that is the context for the command. _nHwndContext_ could be 0.

For  **ExitScope** , _bErrOrCancelled_ is non-zero if the operation failed or was canceled.


## Example

This example shows how to use the  **ExitScope** event. The example determines whether a call to a procedure that handles the **CellChanged** event is in a particular scope?that is, whether the call occurs between the **EnterScope** and **ExitScope** events for that scope.


```vb
 
Private WithEvents vsoApplication As Visio.Application 
Private lngScopeID AsLong 
 
Public Sub Scope_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 'Set the module-level application variable to 
 'trap application-level events. 
 Set vsoApplication = Application 
 
 'Begin a scope. 
 lngScopeID = Application.BeginUndoScope("Draw Shapes") 
 
 'Draw three shapes. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 ActivePage.DrawOval 3, 4, 4, 3 
 ActivePage.DrawLine 4, 5, 5, 4 
 
 'Change a cell (to trigger a CellChanged event). 
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


