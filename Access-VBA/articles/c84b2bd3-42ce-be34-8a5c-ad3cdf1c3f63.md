
# References.ItemAdded Event (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


The  **ItemAdded** event occurs when a reference is added to the project from Visual Basic.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ItemAdded**( **_Reference_**, )

 _expression_A variable that represents a  **References** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Reference|Required| **Reference**|The reference that was added to the project.|

## Remarks
<a name="sectionSection1"> </a>


- The  **ItemAdded** event applies to the ** [References](ac020382-4ece-f138-d1b9-d05b0fe0f523.md)**collection. It isn't associated with a control, form, or report, as are most other events. Therefore, in order to create a procedure definition for the  **ItemAdded** event procedure, you must use a special syntax.
    
- The  **ItemAdded** event can run only an event procedure when it occurs, it cannot run a macro.
    
This event occurs only when you add a reference from code. It doesn't occur when you add a reference from the  **References** dialog box, available by clicking **References** on the **Tools** menu when the Module window is the active window.


## Example
<a name="sectionSection2"> </a>

The following example includes event procedures for the  **ItemAdded** and **ItemRemoved** events. To try this example, first create a new class module by clicking **Class Module** on the **Insert** menu. Paste the following code into the class module and save the module as RefEvents:


```
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
 MsgBox "Reference to " &amp; Reference.Name &amp; " added." 
End Sub 
 
' Display message when reference is removed. 
Private Sub evtReferences_ItemRemoved(ByVal Reference As _ 
 Access.Reference) 
 MsgBox "Reference to " &amp; Reference.Name &amp; " removed." 
End Sub
```

The following Function procedure adds a specified reference. When a reference is added, the ItemAdded event procedure defined in the RefEvents class runs.




```
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
 MsgBox Err &amp; ": " &amp; Err.Description 
 AddReference = False 
 Resume Exit_AddReference 
End Function
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [References Collection](ac020382-4ece-f138-d1b9-d05b0fe0f523.md)
#### Other resources


 [References Object Members](de4ddd41-b41c-6a80-a29c-c2b32d54709a.md)
