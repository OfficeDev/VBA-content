---
title: Application.ProjectBeforeAssignmentNew2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeAssignmentNew2
ms.assetid: 9e2f3358-325e-53b9-3da6-5323482e2a47
ms.date: 06/08/2017
---


# Application.ProjectBeforeAssignmentNew2 Event (Project)

Occurs before one or more assignments are created. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeAssignmentNew2**( ** _pj_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which an assignment or assignments are being created.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the new assignment(s) are not created.|

### Return Value

nothing


## Remarks

The  **ProjectBeforeAssignmentNew2** event also fires when a resource assignment is replaced. Additionally, the event will fire when the only resource assignment on a task is removed, because an "Unassigned Resource" assignment is created after the existing assignment is removed.

Project events do not occur when the project is embedded in another document or application. 

The  **ProjectBeforeAssignmentNew2** event doesn't occur when an assignment is created as the result of a drag-and-drop operation in the **Resource Usage** view, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.


