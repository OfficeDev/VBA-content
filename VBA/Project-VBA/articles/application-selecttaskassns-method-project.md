---
title: Application.SelectTaskAssns Method (Project)
keywords: vbapj.chm1511
f1_keywords:
- vbapj.chm1511
ms.prod: project-server
api_name:
- Project.Application.SelectTaskAssns
ms.assetid: 80683610-657f-f298-0275-831da215a93a
ms.date: 06/08/2017
---


# Application.SelectTaskAssns Method (Project)

Selects all assignments for a selected task in the Team Planner view.


## Syntax

 _expression_. **SelectTaskAssns**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

If a task is assigned to more than one resource, and one assignment is selected, the  **SelectTaskAssns** method selects all assignments in the Team Planner view.


## Example

In the following example, if one task assignment is selected in the Resource Usage view, the view switches to the Team Planner where all assignments for that task are selected.


```vb
Sub SelectAssignments() 
    Application.ViewApply Name:="Team Planner" 
 
    Application.SelectTaskAssns 
End Sub
```


