
# SharedWorkspace.Tasks Property (Office)

 **Last modified:** July 28, 2015

Gets a  ** [SharedWorkspaceTasks](de26341f-44d1-131e-1dbe-e31f3f68e312.md)** collection that represents the list of tasks in the current shared workspace. Read-only.

 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Tasks**

 _expression_A variable that represents a  **SharedWorkspace** object.


## Example

The following example lists the tasks in the current shared workspace.


```
   Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " &amp; swsTasks.Count &amp; _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## See also


#### Concepts


 [SharedWorkspace Object](7512f0ff-382d-d344-9424-aa10549d14f9.md)
#### Other resources


 [SharedWorkspace Object Members](e4c2b518-d955-27e1-3e73-173d3c4f961d.md)
