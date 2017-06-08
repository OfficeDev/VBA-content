---
title: Project.GetTaskIndexByGuid Method (Project)
ms.prod: project-server
api_name:
- Project.Project.GetTaskIndexByGuid
ms.assetid: 6887241c-9daf-385b-42a2-7a82b37c8da7
ms.date: 06/08/2017
---


# Project.GetTaskIndexByGuid Method (Project)

Returns the local task identification number (ID) for the specified task.


## Syntax

 _expression_. **GetTaskIndexByGuid**( ** _TaskGuid_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskGuid_|Required|**String**|The GUID of the task.|

### Return Value

 **Long**


## Remarks

The local task ID is the task index, which changes if the order of the task changes.


## Example

If the ID of the specified task is 6, the following function returns the value 6.


```vb
Function TestTaskId() As Long 
 TestTaskId = ActiveProject.GetTaskIndexByGuid("341A479D-73A5-4209-9366-8EA2B836255B") 
End Function
```


