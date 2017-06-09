---
title: Application.ProjectBeforeAssignmentChange2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeAssignmentChange2
ms.assetid: 99fce7af-00de-42d8-4b61-e97774cc19ed
ms.date: 06/08/2017
---


# Application.ProjectBeforeAssignmentChange2 Event (Project)

Occurs before the user changes the value of an assignment field. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeAssignmentChange2**( ** _asg_**, ** _Field_**, ** _NewVal_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _asg_|Required|**Assignment**|The assignment whose field is being changed.|
| _Field_|Required|**PjAssignmentField**|The field being changed. If more than one field is changed by the user, the event is fired for each field changed. Can be one of the  **[PjAssignmentField](pjassignmentfield-enumeration-project.md)** constants.|
| _NewVal_|Required|**Variant**|The new value for the field specified with Field.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with Field is not changed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application. For more information and sample code for creating and testing an event handler, see [Using Events with Application and Project Objects](using-events-with-application-and-project-objects.md) .

The  **ProjectBeforeAssignmentChange2** event doesn't occur when timescaled data changes, when an entire resource or task row is pasted, when an assignment is changed as the result of a drag-and-drop operation in the **Resource Usage** view, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.


