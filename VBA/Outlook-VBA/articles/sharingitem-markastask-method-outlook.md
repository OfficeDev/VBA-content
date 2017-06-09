---
title: SharingItem.MarkAsTask Method (Outlook)
keywords: vbaol11.chm3223
f1_keywords:
- vbaol11.chm3223
ms.prod: outlook
api_name:
- Outlook.SharingItem.MarkAsTask
ms.assetid: deab1b6c-2d22-678c-1a13-2b171d27a971
ms.date: 06/08/2017
---


# SharingItem.MarkAsTask Method (Outlook)

Marks a  **[SharingItem](sharingitem-object-outlook.md)** object as a task and assigns a task interval for the object.


## Syntax

 _expression_ . **MarkAsTask**( **_MarkInterval_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](olmarkinterval-enumeration-outlook.md)**|The task interval for the  **SharingItem** .|

## Remarks

Calling this method sets the  **[IsMarkedAsTask](sharingitem-ismarkedastask-property-outlook.md)** property to **True** and updates the **[TaskStartDate](sharingitem-taskstartdate-property-outlook.md)** , **[TaskDueDate](sharingitem-taskduedate-property-outlook.md)** , and **[TaskOrdinal](sharingitem-todotaskordinal-property-outlook.md)** properties depending on the value provided in _MarkInterval_.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

