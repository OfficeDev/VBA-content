---
title: TaskRequestAcceptItem.GetAssociatedTask Method (Outlook)
keywords: vbaol11.chm1808
f1_keywords:
- vbaol11.chm1808
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.GetAssociatedTask
ms.assetid: 979459e5-3f95-2e55-f5c9-92e36fc47d5d
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.GetAssociatedTask Method (Outlook)

Returns a  **[TaskItem](taskitem-object-outlook.md)** object that represents the requested task.


## Syntax

 _expression_ . **GetAssociatedTask**( **_AddToTaskList_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AddToTaskList_|Required| **Boolean**| **True** if the task is added to the default **Tasks** folder.|

### Return Value

A  **TaskItem** object that represents the requested task.


## Remarks

The  **GetAssociatedTask** method will not work unless the **TaskItem** is processed before the method is called. To do so, call the **[Display](taskitem-display-method-outlook.md)** method before calling **GetAssociatedTask** .


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

