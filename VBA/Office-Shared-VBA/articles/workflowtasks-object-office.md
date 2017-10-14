---
title: WorkflowTasks Object (Office)
keywords: vbaof11.chm281000
f1_keywords:
- vbaof11.chm281000
ms.prod: office
api_name:
- Office.WorkflowTasks
ms.assetid: 3b0006db-9bad-2dce-d4b1-c67fe5ac54f9
ms.date: 06/08/2017
---


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
|[Application](workflowtasks-application-property-office.md)|
|[Count](workflowtasks-count-property-office.md)|
|[Creator](workflowtasks-creator-property-office.md)|
|[Item](workflowtasks-item-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
