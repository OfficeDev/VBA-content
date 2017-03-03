---
title: WorkflowTask Object (Office)
keywords: vbaof11.chm280000
f1_keywords:
- vbaof11.chm280000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.WorkflowTask
ms.assetid: 9d17947e-f12a-2f97-7888-8d5ec9f85011
---


# WorkflowTask Object (Office)

Represents a single workflow task in a  **WorkflowTasks** collection.


## Example

The following example displays the name of each workflow task in the current document and then displays the workflow task edit user interface for a specific task. It should be noted that calling the  **GetWorkflowTasks** method involves a round-trip to the server.


```vb
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


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

