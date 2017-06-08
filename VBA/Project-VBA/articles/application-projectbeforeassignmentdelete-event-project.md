---
title: Application.ProjectBeforeAssignmentDelete Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeAssignmentDelete
ms.assetid: f0db513e-3dec-e9d6-8385-ac0117e8f28e
ms.date: 06/08/2017
---


# Application.ProjectBeforeAssignmentDelete Event (Project)

Occurs before an assignment is removed or replaced.


## Syntax

 _expression_. **ProjectBeforeAssignmentDelete**( ** _asg_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _asg_|Required|**Assignment**| The assignment that is being removed.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the assignment is not removed. If the assignment is being removed because the associated resource has been deleted, Cancel is ignored.|

### Return Value

nothing


## Remarks

The  **ProjectBeforeAssignmentDelete** event also fires when assigning a resource to a task with no resource assignments, because an "Unassigned Resource" assignment is removed before the new assignment is created.

Project events do not occur when the project is embedded in another document or application. 

The  **ProjectBeforeAssignmentDelete** event doesn't occur when an assignment is deleted as the result of a drag-and-drop operation in the **Resource Usage** view, or when changes have been made using a custom form.


