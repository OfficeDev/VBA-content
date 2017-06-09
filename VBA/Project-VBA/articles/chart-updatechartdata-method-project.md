---
title: Chart.UpdateChartData Method (Project)
keywords: vbapj.chm131637
f1_keywords:
- vbapj.chm131637
ms.prod: project-server
ms.assetid: ecdef74d-480c-05a7-757c-a5c2e3e7359c
ms.date: 06/08/2017
---


# Chart.UpdateChartData Method (Project)
Updates the specified Project data on a chart.

## Syntax

 _expression_. **UpdateChartData** _(Task,_? _Timephased,_? _GroupName,_? _FilterName,_? _LabelField,_? _OutlineLevel,_? _SafeArrayOfPjField,_? _SafeArrayOfPjTimescaledData,_? _TimeScaleUnit,_? _TimescaleUnitCount,_? _StartDate,_? _FinishDate)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Task_|Required|**Boolean**|**True** to update the task data; otherwise, **False**.|
| _Timephased_|Required|**Boolean**|**True** to update the timephased data; otherwise, **False**.|
| _GroupName_|Optional|**String**|The name of the  **[Group2](group2-object-project.md)** object (a group of tasks or resources) for the update.|
| _FilterName_|Optional|**String**|The name of the  **[Filter](filter-object-project.md)** object for the update.|
| _LabelField_|Optional|**PjField**|Specifies the field for the update. Can be one of the  **[PjField](pjfield-enumeration-project.md)** constants.|
| _OutlineLevel_|Optional|**Integer**|Specifies the task outline level for the update. The default value is -1, which is all outline levels.|
| _SafeArrayOfPjField_|Optional|**Variant**|Specifies an array of fields for the update, where each item in the array can be a  **[PjField](pjfield-enumeration-project.md)** constant.|
| _SafeArrayOfPjTimescaledData_|Optional|**Variant**|Specifies an array of timescaled data for the update, where each item in the array can be a  **[PjTimescaledData](pjtimescaleddata-enumeration-project.md)** constant.|
| _TimeScaleUnit_|Optional|**PjTimescaleUnit**|Specifies a timescale unit for the update. Can be a  **[PjTimescaledUnit](pjtimescaleunit-enumeration-project.md)** constant. The default value is **pjTimescaleDays**.|
| _TimescaleUnitCount_|Optional|**Long**|Specifies the number of timescale units to be included in the update. The default value is 1. For example, if the unit is  **pjTimescaleWeeks**, a value of 5 indicates five weeks.|
| _StartDate_|Optional|**Variant**|Specifies the start date for the update.|
| _FinishDate_|Optional|**Variant**|Specifies the finish date for the update.|
| _Task_|Required|BOOL||
| _Timephased_|Required|BOOL||
| _GroupName_|Optional|STRING||
| _FilterName_|Optional|STRING||
| _LabelField_|Optional|PJFIELD||
| _OutlineLevel_|Optional|INT||
| _SafeArrayOfPjField_|Optional|VARIANT||
| _SafeArrayOfPjTimescaledData_|Optional|VARIANT||
| _TimeScaleUnit_|Optional|PJTIMESCALEUNIT||
| _TimescaleUnitCount_|Optional|INT||
| _StartDate_|Optional|VARIANT||
| _FinishDate_|Optional|VARIANT||

### Return value

 **Nothing**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
