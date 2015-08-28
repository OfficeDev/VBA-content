
# StartDriver.OverAllocatedAssignments Property (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets overallocated assignments for a task start driver. Read-only  **OverAllocatedAssignments**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **OverAllocatedAssignments**( **_fOverPeak_**)

 _expression_An expression that returns a  **StartDriver** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|overallocationType|Required| **PjOverallocationType**|Can be one of the  ** [PjOverallocationType](b2eaea51-6884-194c-9a68-75669fcc8283.md)** constants, which determines the type of overallocation.|

## Remarks
<a name="sectionSection1"> </a>

Overallocated assignments are not possible on milestones, placeholder tasks, or tasks with no assignments.


## Example
<a name="sectionSection2"> </a>

The following command returns the number of overallocated assignments where resources are working on other tasks.


```
Debug.Print ActiveProject.Tasks(2).StartDriver.OverAllocatedAssignments(pjOverallocationTypeWorkingOnOtherTasks).Count
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [StartDriver Object](4df2c386-a31e-faea-e286-d510f11cca57.md)
