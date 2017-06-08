---
title: Application.ReminderSet Method (Project)
keywords: vbapj.chm2383
f1_keywords:
- vbapj.chm2383
ms.prod: project-server
api_name:
- Project.Application.ReminderSet
ms.assetid: 5e9305ad-ae42-14e9-8e20-f3068d994200
ms.date: 06/08/2017
---


# Application.ReminderSet Method (Project)

Sets a reminder in Microsoft Outlook for the start time or finish time of the active tasks.


## Syntax

 _expression_. **ReminderSet**( ** _Start_**, ** _LeadTime_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Boolean**|**True** if the reminder is set for the start time of the active tasks. **False** if the reminder is set for the finish time. The default value is **True**.|
| _LeadTime_|Optional|**String**|The amount of lead time for Microsoft Outlook reminders. The default value is "15m", which triggers reminders 15 minutes before the start time (Start is  **True** ) or after the finish time (Start is **False** ).|

### Return Value

 **Boolean**


## Remarks

The  **ReminderSet** method is available only in Project Professional.


