---
title: WorkflowTask.Show Method (Office)
keywords: vbaof11.chm280010
f1_keywords:
- vbaof11.chm280010
ms.prod: office
api_name:
- Office.WorkflowTask.Show
ms.assetid: a7256356-c935-e9ce-e510-6798ebd5563f
ms.date: 06/08/2017
---


# WorkflowTask.Show Method (Office)

Displays a workflow task edit user interface for the specified  **WorkflowTask** object.


## Syntax

 _expression_. **Show**

 _expression_ An expression that returns a **WorkflowTask** object.


### Return Value

Integer


## Example

The following example displays the name of each workflow task in the current document and then displays the workflow task edit user interface for a specific task.


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


[WorkflowTask Object](workflowtask-object-office.md)
#### Other resources


[WorkflowTask Object Members](workflowtask-members-office.md)

