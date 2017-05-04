---
title: WorkflowTasks Object (Office)
keywords: vbaof11.chm281000
f1_keywords:
- vbaof11.chm281000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.WorkflowTasks
ms.assetid: 3b0006db-9bad-2dce-d4b1-c67fe5ac54f9
---


# WorkflowTasks Object (Office)

Represents a collection of  **WorkflowTask** objects.


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

