---
title: Application.ProjectBeforeAssignmentChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeAssignmentChange
ms.assetid: 9d94303c-f8f6-1681-0829-23f240afc570
ms.date: 06/08/2017
---


# Application.ProjectBeforeAssignmentChange Event (Project)

Occurs before the user changes the value of an assignment field.


## Syntax

 _expression_. **ProjectBeforeAssignmentChange**( ** _asg_**, ** _Field_**, ** _NewVal_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _asg_|Required|**Assignment**|The assignment whose field is being changed.|
| _Field_|Required|**PjAssignmentField**| The field being changed. If more than one field is changed by the user, the event is triggered for each field changed. Can be one of the following **[PjAssignmentField](pjassignmentfield-enumeration-project.md)** constants.|
| _NewVal_|Required|**Variant**|The new value for the field specified with  **Field**.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with **Field** is not changed.|

## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeAssignmentChange** event doesn't occur when timescaled data changes, when an entire resource or task row is pasted, when an assignment is changed as the result of a drag-and-drop operation in the Resource Usage view, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form. For more information and sample code for creating and testing an event handler, see[Using Events with Application and Project Objects](using-events-with-application-and-project-objects.md).


## Example

The following example examines new resource assignments and cancels them if they are for the specified resource. This example requires a new class module and additional code for it to have an effect.


```vb
Private Sub App_ProjectBeforeAssignmentChange(ByVal asg As Assignment, ByVal Field As PjAssignmentField, _ 
    ByVal NewVal As Variant, Cancel As Boolean) 
 
    If Field = pjAssignmentResourceName And NewVal = "Lisa Jones" Then 
        MsgBox "Lisa is no longer available for assignment!" 
        Cancel = True 
    End If 
End Sub
```


