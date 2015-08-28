
# WorkflowTasks Object (Office)

 **Last modified:** July 28, 2015

Represents a collection of  **WorkflowTask** objects.

## Example

The following example displays the name of each workflow task in the current document and then displays the workflow task edit user interface for a specific task. It should be noted that calling the  **GetWorkflowTasks** method involves a round-trip to the server.


```
Sub DisplayWorkTask() 
Dim objWorkflowTasks As WorkflowTasks 
Dim objWorkflowTask As WorkflowTask 
Dim cnt As Integer 
 
Set objWorkflowTasks = Document.GetWorkflowTasks() 
 
For cnt = 1 To objWorkflowTasks.Count 
 Debug.Print objWorkflowTask(cnt).Name 
Next 
 
Set objWorkflowTask = objWorkflowTasks(1) 
objWorkflowTask.Show 
 
End Sub 

```


## See also


#### Concepts


 [Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Other resources


 [WorkflowTasks Object Members](a627f77c-fd47-ef66-edbd-9b4c4fcd9920.md)
