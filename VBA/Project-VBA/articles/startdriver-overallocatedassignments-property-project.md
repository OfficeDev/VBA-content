---
title: StartDriver.OverAllocatedAssignments Property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.OverAllocatedAssignments
ms.assetid: bef55fa0-e721-27f6-aa3b-6314aeaef0fa
ms.date: 06/08/2017
---


# StartDriver.OverAllocatedAssignments Property (Project)

Gets overallocated assignments for a task start driver. Read-only  **OverAllocatedAssignments**.


## Syntax

 _expression_. **OverAllocatedAssignments**( ** _fOverPeak_** )

 _expression_ An expression that returns a **StartDriver** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _overallocationType_|Required|**PjOverallocationType**|Can be one of the  **[PjOverallocationType](pjoverallocationtype-enumeration-project.md)** constants, which determines the type of overallocation.|

## Remarks

Overallocated assignments are not possible on milestones, placeholder tasks, or tasks with no assignments.


## Example

The following command returns the number of overallocated assignments where resources are working on other tasks.


```vb
Debug.Print ActiveProject.Tasks(2).StartDriver.OverAllocatedAssignments(pjOverallocationTypeWorkingOnOtherTasks).Count
```


## See also


#### Concepts


[StartDriver Object](startdriver-object-project.md)
