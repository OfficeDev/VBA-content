
# WorkflowTasks Object (Office)

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


## Properties



|**Name**|
|:-----|
|[Application](65ecee81-f689-a72e-6b77-91142dcbfe18.md)|
|[Count](0c1dafe0-d89e-d7b4-1461-5c78db47cae9.md)|
|[Creator](9554018d-322d-dc5d-787a-c0b0e9f9da44.md)|
|[Item](f47adb68-5cfb-c3d0-e887-5a6d587a51b3.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)