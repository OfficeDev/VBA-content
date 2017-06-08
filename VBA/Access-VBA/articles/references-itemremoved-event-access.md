---
title: References.ItemRemoved Event (Access)
keywords: vbaac10.chm12647
f1_keywords:
- vbaac10.chm12647
ms.prod: access
api_name:
- Access.References.ItemRemoved
ms.assetid: 19498b96-5e92-8a7a-512a-95a89b878eb2
ms.date: 06/08/2017
---


# References.ItemRemoved Event (Access)

The  **ItemRemoved** event occurs when a reference is removed from the project.


## Syntax

 _expression_. **ItemRemoved**( ** _Reference_**, )

 _expression_ A variable that represents a **References** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Reference_|Required|**Reference**|The reference that was removed from the project.|

## Remarks


- The  **ItemRemoved** event applies to the **[References](references-object-access.md)** collection. It isn't associated with a control, form, or report, as are most other events. Therefore, in order to create a procedure definition for the **ItemRemoved** event procedure, you must use a special syntax.
    
- The  **ItemRemoved** event can run only an event procedure when it occurs, it cannot run a macro.
    
This event occurs only when you remove a reference from code. It doesn't occur when you remove a reference from the  **References** dialog box, available by clicking **References** on the **Tools** menu when the Module window is the active window.


## Example

The following example includes event procedures for the  **ItemAdded** and **ItemRemoved** events. To try this example, first create a new class module by clicking **Class Module** on the **Insert** menu. Paste the following code into the class module and save the module as RefEvents:


```vb
' Declare object variable to represent References collection. 
Public WithEvents evtReferences As References 
 
' When instance of class is created, initialize evtReferences 
' variable. 
Private Sub Class_Initialize() 
 Set evtReferences = Application.References 
End Sub 
 
' When instance is removed, set evtReferences to Nothing. 
Private Sub Class_Terminate() 
 Set evtReferences = Nothing 
End Sub 
 
' Display message when reference is added. 
Private Sub evtReferences_ItemAdded(ByVal Reference As _ 
 Access.Reference) 
 MsgBox "Reference to " &; Reference.Name &; " added." 
End Sub 
 
' Display message when reference is removed. 
Private Sub evtReferences_ItemRemoved(ByVal Reference As _ 
 Access.Reference) 
 MsgBox "Reference to " &; Reference.Name &; " removed." 
End Sub
```

The next Function procedure removes a specified reference. When a reference is removed, the ItemRemoved event procedure defined in the RefEvents class runs.

For example, to remove a reference to the calendar control, you could pass the string "MSACAL", which is the name of the  **Reference** object that represents the calendar control.




```vb
Function RemoveReference(strRefName As String) As Boolean 
 Dim ref As Reference 
 
 On Error GoTo Error_RemoveReference 
 ' Return object representing existing reference. 
 Set ref = objRefEvents.evtReferences(strRefName) 
 ' Remove reference from collection. 
 objRefEvents.evtReferences.Remove ref 
 RemoveReference = True 
 
Exit_RemoveReference: 
 Exit Function 
 
Error_RemoveReference: 
 MsgBox Err &; ": " &; Err.Description 
 RemoveReference = False 
 Resume Exit_RemoveReference 
End Function
```


## See also


#### Concepts


[References Collection](references-object-access.md)

