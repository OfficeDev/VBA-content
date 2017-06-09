---
title: Assignment.Replan Method (Project)
keywords: vbapj.chm131251
f1_keywords:
- vbapj.chm131251
ms.prod: project-server
api_name:
- Project.Assignment.Replan
ms.assetid: 29ec0102-b4e4-c9dc-d930-4f8ff4069bd6
ms.date: 06/08/2017
---


# Assignment.Replan Method (Project)

Replans the assignment by decreasing work or increasing duration.


## Syntax

 _expression_. **Replan**( ** _action_** )

 _expression_ An expression that returns a **Assignment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _action_|Required|**PjAssignmentReplanAction**|Can be one of the following constants in  **[PjAssignmentReplanAction](pjassignmentreplanaction-enumeration-project.md)**: **pjConstrainToMaxUnitsByDecreasingWork** or **pjConstrainToMaxUnitsByIncreasingDuration**.|

### Return Value

Nothing


## Remarks

For example, if a resource calendar changes so that the resource becomes overallocated, you can replan the overallocated assignments.


## Example

In the following example, an overallocated assignment selected in the Team Planner view is changed to increased duration.


```vb
ActiveCell.Assignment.Replan(pjConstrainToMaxUnitsByIncreasingDuration)
```


