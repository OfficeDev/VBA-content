---
title: TaskRequestDeclineItem.GetAssociatedTask Method (Outlook)
keywords: vbaol11.chm1857
f1_keywords:
- vbaol11.chm1857
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.GetAssociatedTask
ms.assetid: 4d92d092-b8b8-4378-1193-8b7f17b9dacb
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.GetAssociatedTask Method (Outlook)

Returns a  **[TaskItem](taskitem-object-outlook.md)** object that represents the requested task.


## Syntax

 _expression_ . **GetAssociatedTask**( **_AddToTaskList_** )

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


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


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

