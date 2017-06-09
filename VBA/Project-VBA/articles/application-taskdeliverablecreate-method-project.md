---
title: Application.TaskDeliverableCreate Method (Project)
keywords: vbapj.chm92
f1_keywords:
- vbapj.chm92
ms.prod: project-server
api_name:
- Project.Application.TaskDeliverableCreate
ms.assetid: 61bd8608-8a5f-3555-b769-5ee951f8ebd7
ms.date: 06/08/2017
---


# Application.TaskDeliverableCreate Method (Project)

Creates or removes a deliverable for the selected task. Available only in Project Professional.


## Syntax

 _expression_. **TaskDeliverableCreate**( ** _Create_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Create_|Optional|**Variant**|If the selected task has no associated deliverable,  **True** creates a deliverable. If the selected task does have an associated deliverable, **False** removes the deliverable. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

When the selected task does not have a deliverable, following are results of running the  **TaskDeliverableCreate** method:


-  `TaskDeliverableCreate(True)` creates a deliverable for the selected task.
    
-  `TaskDeliverableCreate(False)` does nothing.
    


When the selected task has an associated deliverable, following are results of running the  **TaskDeliverableCreate** method:


-  `TaskDeliverableCreate(True)` gives the error, **Cannot create a deliverable link for the selected subproject task.**, followed by the run-time error 1004, **An unexpected error occurred with the method.**
    
-  `TaskDeliverableCreate(False)` removes the deliverable.
    


The  **TaskDeliverableCreate** method is equivalent to the **Create Deliverables** command on the **Deliverable** drop-down menu on the **Task** tab of the Ribbon. If the selected task has no deliverable, the **Create Deliverables** command creates one. If the selected task has a deliverable, **Create Deliverables** shows an active icon, and selecting the command deletes the deliverable.


 **Note**  You cannot create a task deliverable until you publish the project and create a project workspace. You also cannot create a deliverable on a summary task.


## Example

The following example creates or deletes a deliverable for the selected task in a published project.


```vb
Sub ToggleDeliverable() 
    Dim deliverGuid As String 
 
    deliverGuid = ActiveCell.Task.deliverableGuid 
 
    If deliverGuid = "00000000-0000-0000-0000-000000000000" Then 
        TaskDeliverableCreate Create:=True 
    Else 
        TaskDeliverableCreate Create:=False 
    End If 
End Sub
```


