---
title: Application.ProjectBeforeAssignmentDelete2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeAssignmentDelete2
ms.assetid: 2753a140-e01b-b2c1-233f-f9f265737b47
ms.date: 06/08/2017
---


# Application.ProjectBeforeAssignmentDelete2 Event (Project)

Occurs before an assignment is removed or replaced. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeAssignmentDelete2**( ** _asg_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _asg_|Required|**Assignment**|The assignment that is being removed.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the assignment is not removed. If the assignment is being removed because the associated resource has been deleted, Info is ignored.|

### Return Value

nothing


## Remarks

The  **ProjectBeforeAssignmentDelete2** event also fires when assigning a resource to a task with no resource assignments, because an "Unassigned Resource" assignment is removed before the new assignment is created.

Project events do not occur when the project is embedded in another document or application. 

The  **ProjectBeforeAssignmentDelete2** event doesn't occur when an assignment is deleted as the result of a drag-and-drop operation in the **Resource Usage** view, or when changes have been made using a custom form.


