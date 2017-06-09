---
title: Application.UpdateTasks Method (Project)
keywords: vbapj.chm2350
f1_keywords:
- vbapj.chm2350
ms.prod: project-server
api_name:
- Project.Application.UpdateTasks
ms.assetid: 4a04e459-9f5c-f944-d39f-dcbbfc48fdab
ms.date: 06/08/2017
---


# Application.UpdateTasks Method (Project)

Updates the selected tasks.


## Syntax

 _expression_. **UpdateTasks**( ** _PercentComplete_**, ** _ActualDuration_**, ** _RemainingDuration_**, ** _ActualStart_**, ** _ActualFinish_**, ** _Notes_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PercentComplete_|Optional|**Variant**|The percent complete of the active tasks.|
| _ActualDuration_|Optional|**Variant**|The actual duration of the selected tasks.|
| _RemainingDuration_|Optional|**Variant**|The remaining duration of the selected tasks.|
| _ActualStart_|Optional|**Variant**|The actual start date of the selected tasks.|
| _ActualFinish_|Optional|**Variant**|The actual finish date of the selected tasks.|
| _Notes_|Optional|**String**|Comments in the Notes field for the selected tasks. The value can be text only, not Rich Text Format (RTF) as in the  **Notes** dialog box.|

### Return Value

 **Boolean**


## Remarks

Using the  **UpdateTasks** method without specifying any arguments displays the **Update Tasks** dialog box.


## Example

The following example creates a task named "TestTask-1", updates the task to 50% complete, and then deletes the task. 


```vb
Sub Update_Tasks() 
 
 'Activate Gantt Chart 
 ViewApply Name:="Gantt Chart" 
 
 'Create a task 
 RowInsert 
 SetTaskField Field:="Name", Value:="TestTask-1" 
 SetTaskField Field:="Duration", Value:="2" 
 
 'Update the percent complete of the new task. 
 UpdateTasks PercentComplete:="50" 
 
 'Delete the new task 
 ActiveProject.Tasks("TestTask-1").Delete 
End Sub
```


