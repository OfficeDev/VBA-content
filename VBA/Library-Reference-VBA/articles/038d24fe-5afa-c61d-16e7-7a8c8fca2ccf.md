
# SharedWorkspaceTask.Title Property (Office)

Sets or gets the title of a  **SharedWorkspaceTask** object. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Title**

 _expression_ A variable that represents a **SharedWorkspaceTask** object.


### Return Value

String


## Remarks

The  **Title** property is the single required property of a shared workspace task. Use the optional **Description** property to provide or return additional information about the task.


## Example

The following example displays a list of the titles of all tasks in the current shared workspace.


```vb
 Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTasks As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        strTasks = strTasks &; swsTask.Title &; vbCrLf 
    Next 
    MsgBox strTasks, vbInformation + vbOKOnly, _ 
        "Tasks in Shared Workspace" 
    Set swsTask = Nothing 
 

```


## See also


#### Concepts


[SharedWorkspaceTask Object](fbd82b03-53fa-12ff-9fb2-07bef012dde8.md)
#### Other resources


[SharedWorkspaceTask Object Members](5b5589d1-f907-7357-f930-eede569d2021.md)