
# StartDriver.OverAllocatedAssignments Property (Project)

Gets overallocated assignments for a task start driver. Read-only  **OverAllocatedAssignments**.


## Syntax

 _expression_. **OverAllocatedAssignments**( ** _fOverPeak_** )

 _expression_ An expression that returns a **StartDriver** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _overallocationType_|Required|**PjOverallocationType**|Can be one of the  **[PjOverallocationType](b2eaea51-6884-194c-9a68-75669fcc8283.md)** constants, which determines the type of overallocation.|

## Remarks

Overallocated assignments are not possible on milestones, placeholder tasks, or tasks with no assignments.


## Example

The following command returns the number of overallocated assignments where resources are working on other tasks.


```vb
Debug.Print ActiveProject.Tasks(2).StartDriver.OverAllocatedAssignments(pjOverallocationTypeWorkingOnOtherTasks).Count
```


## See also


#### Concepts


[StartDriver Object](4df2c386-a31e-faea-e286-d510f11cca57.md)