---
title: Application.UpdateProject Method (Project)
keywords: vbapj.chm611
f1_keywords:
- vbapj.chm611
ms.prod: project-server
api_name:
- Project.Application.UpdateProject
ms.assetid: a6f80334-7faf-ca95-b5ed-0a9fba516169
ms.date: 06/08/2017
---


# Application.UpdateProject Method (Project)

Updates progress information and reschedules work for tasks in a project.


## Syntax

 _expression_. **UpdateProject**( ** _All_**, ** _UpdateDate_**, ** _Action_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _All_|Optional|**Boolean**|**True** if all tasks in the active project are updated. **False** if only the selected tasks are updated. The default value is **True**.|
| _UpdateDate_|Optional|**Variant**|The update date to use for the specified action. |
| _action_|Optional|**Integer**|The action to take with the specified tasks. Can be one of the following  **[PjProjectUpdate](pjprojectupdate-enumeration-project.md)** constants: **pj0or100Percent**, **pj0to100Percent**, or **pjReschedule**. The default is **pj0to100Percent**.|

### Return Value

 **Boolean**


## Remarks

Running the  **UpdateProject** method with no arguments displays the **Update Project** dialog box.

The  **UpdateProject** method corresponds to the **Update Project** command on the **PROJECT** tab of the ribbon.


## Example

The following example first creates a task, sets the "% Complete" field to 50 percent; and then updates the project to schedule the rest of the work for the task to start on 9/19/2012.


```vb
Sub Update_Project() 
    ViewApply Name:="Gantt Chart" 
 
    ' Create a new task 
    RowInsert 
    SetTaskField Field:="Name", Value:="TestTask-1" 
    SetTaskField Field:="Duration", Value:="2" 
    SetTaskField Field:="% Complete", Value:="50" 
 
    'Schedule the remainder of the work to start on the update date. 
    UpdateProject All:=False, UpdateDate:="9/19/2012", action:=pjReschedule 
End Sub
```


