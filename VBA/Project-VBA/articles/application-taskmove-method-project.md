---
title: Application.TaskMove Method (Project)
keywords: vbapj.chm2289
f1_keywords:
- vbapj.chm2289
ms.prod: project-server
api_name:
- Project.Application.TaskMove
ms.assetid: 7a847c59-b07c-6bf2-90a3-b62d0d080cc6
ms.date: 06/08/2017
---


# Application.TaskMove Method (Project)

Moves the start date of one or more selected tasks the specified number of days.


## Syntax

 _expression_. **TaskMove**( ** _MoveForward_**, ** _IsWorkingDuration_**, ** _MoveDays_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MoveForward_|Optional|**Variant**|**True** if the task moves forward in time. **False** if the task moves backward in time. The default is **True**.|
| _IsWorkingDuration_|Optional|**Variant**|**True** if the the number of days specified by _MoveDays_ is only for working days. **False** if the number of days specified by _MoveDays_ includes both working and nonworking days. The default is **True**.|
| _MoveDays_|Optional|**Integer**|Specifies the number of days to move the selected tasks. The default value is 1.|

### Return Value

 **Boolean**


## Remarks

The  **TaskMove** method does not override a predecessor task constraint for automatically scheduled tasks.

The  **TaskMove** method corresponds to various commands in the **Move Task** drop-down menu on the **TASK** ribbon. To move incomplete or complete parts of a task to the status date, use the **[TaskMoveToStatusDate](application-taskmovetostatusdate-method-project.md)** method.


## Example

For the following example, a selected task start date is Friday, 7/24/09. After running the statement, the start date of the task is Monday, 8/3/09. The start date of the task has moved forward eight working days.


```vb
Application.TaskMove MoveDays:=8
```

If the selected task is manually scheduled and has a predecessor task with a finish-to-start (FS) constraint, the following statement moves the selected task back one working day.

If you change the selected task to automatically scheduled, the statement can move the task back only as far as the finish date of the predecessor task.




```vb
Application.TaskMove MoveForward:=False
```


