---
title: Application.OptionsSchedule Method (Project)
keywords: vbapj.chm644
f1_keywords:
- vbapj.chm644
ms.prod: project-server
api_name:
- Project.Application.OptionsSchedule
ms.assetid: 24035b34-0364-e830-864a-801150e2668d
ms.date: 06/08/2017
---


# Application.OptionsSchedule Method (Project)

Sets scheduling options.


## Syntax

 _expression_. **OptionsSchedule**( ** _ScheduleMessages_**, ** _StartOnCurrentDate_**, ** _AutoLink_**, ** _AutoSplit_**, ** _CriticalSlack_**, ** _TaskType_**, ** _DurationUnits_**, ** _WorkUnits_**, ** _AutoTrack_**, ** _SetDefaults_**, ** _AssignmentUnits_**, ** _EffortDriven_**, ** _HonorConstraints_**, ** _ShowEstimated_**, ** _NewTasksEstimated_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ScheduleMessages_|Optional|**Variant**|**Boolean**. **True** if messages display when scheduling problems occur; otherwise, **False**.|
| _StartOnCurrentDate_|Optional|**Boolean**|**True** if new tasks start on the current date. **False** if new tasks start on the project start date (projects scheduled from the start date) or on the project finish date (projects scheduled from the finish date).|
| _AutoLink_|Optional|**Boolean**|**True** if tasks are automatically linked; otherwise, **False**.|
| _AutoSplit_|Optional|**Boolean**|**True** if tasks in progress are automatically split; otherwise, **False**.|
| _CriticalSlack_|Optional|**Variant**|The maximum amount of slack allowed for critical tasks.|
| _TaskType_|Optional|**Long**|The default type for new tasks. Can be one of the  **[PjTaskFixedType](pjtaskfixedtype-enumeration-project.md)** constants.|
| _DurationUnits_|Optional|**Long**|The default duration unit for tasks. Can be one of the  **[PjUnit](pjunit-enumeration-project.md)** constants.|
| _WorkUnits_|Optional|**Long**|The default work unit for resource assignments. Can be one of the  **PjUnit** constants.|
| _AutoTrack_|Optional|**Boolean**|**True** if task tracking fields automatically update resource assignments; otherwise, **False**.|
| _SetDefaults_|Optional|**Boolean**|**True** if the values specified for all arguments except ScheduleMessages and AssignmentUnits become the defaults for new project files; otherwise, **False**.|
| _AssignmentUnits_|Optional|**Long**|Specifies how assignment units should display. Can be one of the  **[PjAssignmentUnit](pjassignmentunits-enumeration-project.md)** constants.|
| _EffortDriven_|Optional|**Boolean**|**True** if new tasks are effort-driven; otherwise, **False**.|
| _HonorConstraints_|Optional|**Boolean**|**True** if tasks honor their constraint dates; otherwise, **False**.|
| _ShowEstimated_|Optional|**Boolean**|**True** if task durations in new projects are displayed with the estimated character; otherwise, **False**.|
| _NewTasksEstimated_|Optional|**Boolean**|**True** if new tasks in the active project have estimated durations; otherwise, **False**.|

### Return Value

Boolean


## Remarks

If an argument is omitted, its default value is specified by the current setting on the  **Schedule** tab of the **Project Options** dialog box.

Using the  **OptionsSchedule** method without specifying any arguments displays the **Project Options** dialog box.


## Example

The following example enables messages to be displayed when scheduling problems occur, schedules new tasks to start on the current date, and sets the default duration unit for tasks to a week.


```vb
Sub Options_Schedule() 
 OptionsSchedule ScheduleMessages:=True, StartOnCurrentDate:=True, DurationUnits:=pjWeek 
End Sub
```


