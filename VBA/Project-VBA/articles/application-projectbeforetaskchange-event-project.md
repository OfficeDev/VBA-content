---
title: Application.ProjectBeforeTaskChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskChange
ms.assetid: 995024c3-b031-0ddd-0fbe-4d817f237473
ms.date: 06/08/2017
---


# Application.ProjectBeforeTaskChange Event (Project)

Occurs before the user changes the value of a task field.


## Syntax

 _expression_. **ProjectBeforeTaskChange**( ** _tsk_**, ** _Field_**, ** _NewVal_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _tsk_|Required|**Task**|The task whose field is being changed.|
| _Field_|Required|**Long**|The field being changed. If more than one field is changed by the user, the event is fired for each field changed. Can be one of the  **[PjField](pjfield-enumeration-project.md)** constants.|
| _NewVal_|Required|**Variant**|The new value for the field specified with  **Field**.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with **Field** is not changed.|

## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeTaskChange** event doesn't occur when timescaled data changes, when constraint data in the Task Details Form changes, when a task is split by manipulating its task bar on the Gantt Chart, when changes are made to outline level or outline number, when a baseline is saved, when a baseline is cleared, when an entire task row is pasted, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form. For more information and sample code for creating and testing an event handler, see[Using Events with Application and Project Objects](using-events-with-application-and-project-objects.md).


## Example

The following example informs the user when the duration of a task increases and by how much. This example requires a new class module and additional code for it to have an effect.


```vb
Private Sub App_ProjectBeforeTaskChange(ByVal tsk As Task, ByVal Field As PjField, _ 
    ByVal NewVal As Variant, Cancel As Boolean) 
 
    Dim TaskDuration As Long 
 
    TaskDuration = Val(NewVal) * 480 ' Convert days to minutes 
 
    If Field = pjTaskDuration And TaskDuration > tsk.Duration Then 
        If (TaskDuration - tsk.Duration) \ 480 < 1 Then 
            MsgBox "The task " &; Chr$(34) &; tsk.Name &; Chr$(34) &; " is now " &; _ 
                (TaskDuration - tsk.Duration) / 480 &; (TaskDuration - tsk.Duration) \ 480 &; _ 
                " day(s) longer." 
        Else 
            MsgBox "The task " &; Chr$(34) &; tsk.Name &; Chr$(34) &; " is now " &; _ 
               (TaskDuration - tsk.Duration) / 480 &; " day(s) longer." 
        End If 
    End If 
End Sub
```


