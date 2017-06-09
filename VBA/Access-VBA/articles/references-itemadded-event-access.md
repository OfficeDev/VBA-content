---
title: References.ItemAdded Event (Access)
keywords: vbaac10.chm12646
f1_keywords:
- vbaac10.chm12646
ms.prod: access
api_name:
- Access.References.ItemAdded
ms.assetid: c84b2bd3-42ce-be34-8a5c-ad3cdf1c3f63
ms.date: 06/08/2017
---


# References.ItemAdded Event (Access)

The  **ItemAdded** event occurs when a reference is added to the project from Visual Basic.


## Syntax

 _expression_. **ItemAdded**( ** _Reference_**, )

 _expression_ A variable that represents a **References** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Reference_|Required|**Reference**|The reference that was added to the project.|

## Remarks


- The  **ItemAdded** event applies to the **[References](references-object-access.md)** collection. It isn't associated with a control, form, or report, as are most other events. Therefore, in order to create a procedure definition for the **ItemAdded** event procedure, you must use a special syntax.
    
- The  **ItemAdded** event can run only an event procedure when it occurs, it cannot run a macro.
    
This event occurs only when you add a reference from code. It doesn't occur when you add a reference from the  **References** dialog box, available by clicking **References** on the **Tools** menu when the Module window is the active window.


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

The following Function procedure adds a specified reference. When a reference is added, the ItemAdded event procedure defined in the RefEvents class runs.




```vb
' Create new instance of RefEvents class. 
Dim objRefEvents As New RefEvents 
 
' Pass file name and path of type library to this procedure. 
Function AddReference(strFileName As String) As Boolean 
 Dim ref As Reference 
 
 On Error GoTo Error_AddReference 
 ' Create new reference on References object variable. 
 Set ref = objRefEvents.evtReferences.AddFromFile(strFileName) 
 AddReference = True 
 
Exit_AddReference: 
 Exit Function 
 
Error_AddReference: 
 MsgBox Err &; ": " &; Err.Description 
 AddReference = False 
 Resume Exit_AddReference 
End Function
```


## See also


#### Concepts


[References Collection](references-object-access.md)

