---
title: WorkflowTask Object (Office)
keywords: vbaof11.chm280000
f1_keywords:
- vbaof11.chm280000
ms.prod: office
api_name:
- Office.WorkflowTask
ms.assetid: 9d17947e-f12a-2f97-7888-8d5ec9f85011
ms.date: 06/08/2017
---


# WorkflowTask Object (Office)

Represents a single workflow task in a  **WorkflowTasks** collection.


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


## Methods



|**Name**|
|:-----|
|[Show](workflowtask-show-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](workflowtask-application-property-office.md)|
|[AssignedTo](workflowtask-assignedto-property-office.md)|
|[CreatedBy](workflowtask-createdby-property-office.md)|
|[CreatedDate](workflowtask-createddate-property-office.md)|
|[Creator](workflowtask-creator-property-office.md)|
|[Description](workflowtask-description-property-office.md)|
|[DueDate](workflowtask-duedate-property-office.md)|
|[Id](workflowtask-id-property-office.md)|
|[ListID](workflowtask-listid-property-office.md)|
|[Name](workflowtask-name-property-office.md)|
|[WorkflowID](workflowtask-workflowid-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
