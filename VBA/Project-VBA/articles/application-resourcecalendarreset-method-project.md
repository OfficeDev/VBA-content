---
title: Application.ResourceCalendarReset Method (Project)
keywords: vbapj.chm621
f1_keywords:
- vbapj.chm621
ms.prod: project-server
api_name:
- Project.Application.ResourceCalendarReset
ms.assetid: 3dd5a235-c855-0d65-a664-655c9c1fa7b0
ms.date: 06/08/2017
---


# Application.ResourceCalendarReset Method (Project)

Resets a resource calendar.


## Syntax

 _expression_. **ResourceCalendarReset**( ** _ProjectName_**, ** _ResourceName_**, ** _BaseCalendar_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectName_|Required|**String**|The name of the project containing the resource calendar to reset.|
| _ResourceName_|Required|**String**|The name of the resource for the calendar to reset.|
| _BaseCalendar_|Optional|**String**|The name of the base calendar used to reset the resource calendar. The default value is the name of the current base calendar for the resource.|

### Return Value

 **Boolean**


## Remarks

The  **ResourceCalendarReset** method has no effect for material resources.


