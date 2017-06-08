---
title: Application.SetTaskField Method (Project)
keywords: vbapj.chm4
f1_keywords:
- vbapj.chm4
ms.prod: project-server
api_name:
- Project.Application.SetTaskField
ms.assetid: 44e3df27-8924-ecbb-b655-7dab9a51d96f
ms.date: 06/08/2017
---


# Application.SetTaskField Method (Project)

Sets the value of a task field specified by the name of the field.


## Syntax

 _expression_. **SetTaskField**( ** _Field_**, ** _Value_**, ** _AllSelectedTasks_**, ** _Create_**, ** _TaskID_**, ** _ProjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**|The name of the task field to set.|
| _Value_|Required|**String**|The value of the task field.|
| _AllSelectedTasks_|Optional|**Boolean**|**True** if the value of the field is set for all selected tasks. **False** if the value is set for the active task. The default value is **False**.|
| _Create_|Optional|**Boolean**|**True** if Project creates a task when the active cell is on an empty row. The default value is **True**.|
| _TaskID_|Optional|**Long**|The identification number of the task containing the field to set. If  _AllSelectedTasks_ is **True**,  _TaskID_ is ignored.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the task specified by  _TaskID_. If  _TaskID_ is not specified, _ProjectName_ is ignored. The default value is the name of the active project.|

### Return Value

 **Boolean**


## Remarks

To set a task field by ID, use the  **[SetTaskFieldByID](application-settaskfieldbyid-method-project.md)** method.


## Example

The following example changes the task field "Name" of Task ID 3 to "New Task Name", and then changes it back to the original name.


```vb
Sub Set_TaskField() 
    Dim T As Task 
    Set T = ActiveProject.Tasks(3)
 
    ' Save the task name 
    OldName = T.GetField(pjTaskName) 
 
    ViewApply Name:="&;Gantt Chart" 
    SetTaskField Field:="Name", Value:="New Task's Name", TaskID:=3 
    SetTaskField Field:="Name", Value:=OldName, TaskID:=3 
End Sub
```


