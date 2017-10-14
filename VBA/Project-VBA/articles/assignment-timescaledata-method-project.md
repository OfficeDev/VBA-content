---
title: Assignment.TimeScaleData Method (Project)
ms.prod: project-server
api_name:
- Project.Assignment.TimeScaleData
ms.assetid: ff948754-cc0e-8bf0-31e8-30b19dbcb08d
ms.date: 06/08/2017
---


# Assignment.TimeScaleData Method (Project)

Sets options for displaying timephased data.


## Syntax

 _expression_. **TimeScaleData**( ** _StartDate_**, ** _EndDate_**, ** _Type_**, ** _TimeScaleUnit_**, ** _Count_** )

 _expression_ A variable that represents an **Assignment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartDate_|Required|**Variant**|The start date for the timephased data. If the start date falls within an interval, it is "rounded" to the start of the interval. For example, if TimeScaleUnit is  **pjTimescaleWeeks** and StartDate specifies a Wednesday, the start date is rounded to the preceding Monday (assuming that the work week starts on a Monday).|
| _EndDate_|Required|**Variant**|The end date for the timephased data. If the end date falls within an interval, it is "rounded" to the end of the interval.|
| _Type_|Optional|**Long**|The type of timephased data. Can be one of the  **[PjAssignmentTimescaledData ](pjassignmenttimescaleddata-enumeration-project.md)** constants. The default value is **pjAssignmentTimescaledWork**.|
| _TimeScaleUnit_|Optional|**Long**|Can be one of the  **[PjTimescaleUnit](pjtimescaleunit-enumeration-project.md)** constants. The default value is **pjTimescaleWeeks**.|
| _Count_|Optional|**Long**|The number of timescale units to group together. The default value is 1. |

### Return Value

 **TimeScaleValues**


