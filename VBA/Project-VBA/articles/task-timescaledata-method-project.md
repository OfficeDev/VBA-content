---
title: Task.TimeScaleData Method (Project)
ms.prod: project-server
api_name:
- Project.Task.TimeScaleData
ms.assetid: 58526bce-9ee0-8dce-98ee-a8b8e07175eb
ms.date: 06/08/2017
---


# Task.TimeScaleData Method (Project)

Sets options for displaying timephased data for the task.


## Syntax

 _expression_. **TimeScaleData**( ** _StartDate_**, ** _EndDate_**, ** _Type_**, ** _TimeScaleUnit_**, ** _Count_** )

 _expression_ A variable that represents a **Task** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartDate_|Required|**Variant**|The start date for the timephased data. If the start date falls within an interval, it is "rounded" to the start of the interval. For example, if TimeScaleUnit is  **pjTimescaleWeeks** and StartDate specifies a Wednesday, the start date is rounded to the preceding Monday (assuming that the work week started on a Monday).|
| _EndDate_|Required|**Variant**|The end date for the timephased data. If the end date falls within an interval, it is "rounded" to the end of the interval.|
| _Type_|Optional|**Long**|The type of timephased data. Can be one of the  **[PjTaskTimescaledData](pjtasktimescaleddata-enumeration-project.md)** constants. The default value is **pjTaskTimescaledWork**.|
| _TimeScaleUnit_|Optional|**Long**|Can be one of the  **[PjTimescaleUnit](pjtimescaleunit-enumeration-project.md)** constants. The default value is **pjTimescaleWeeks**.|
| _Count_|Optional|**Long**|The number of timescale units to group together. The default value is 1.|

### Return Value

 **TimeScaleValues**


