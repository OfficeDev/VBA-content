
# SharedWorkspaceTask.Title Property (Office)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)
 [Example](#sectionSection3)


Sets or gets the title of a  **SharedWorkspaceTask** object. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **Title**

 _expression_A variable that represents a  **SharedWorkspaceTask** object.


### Return Value

String


## Remarks
<a name="sectionSection2"> </a>

The  **Title** property is the single required property of a shared workspace task. Use the optional **Description** property to provide or return additional information about the task.


## Example
<a name="sectionSection3"> </a>

The following example displays a list of the titles of all tasks in the current shared workspace.


```
 Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTasks As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        strTasks = strTasks &amp; swsTask.Title &amp; vbCrLf 
    Next 
    MsgBox strTasks, vbInformation + vbOKOnly, _ 
        "Tasks in Shared Workspace" 
    Set swsTask = Nothing 
 

```


## See also
<a name="sectionSection3"> </a>


#### Concepts


 [SharedWorkspaceTask Object](fbd82b03-53fa-12ff-9fb2-07bef012dde8.md)
#### Other resources


 [SharedWorkspaceTask Object Members](5b5589d1-f907-7357-f930-eede569d2021.md)
