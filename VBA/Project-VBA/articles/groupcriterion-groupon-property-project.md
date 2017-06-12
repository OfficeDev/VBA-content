---
title: GroupCriterion.GroupOn Property (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion.GroupOn
ms.assetid: dd36cf16-9306-4cc7-904b-9e2ae364722f
ms.date: 06/08/2017
---


# GroupCriterion.GroupOn Property (Project)

Gets or sets the type of grouping for a field used as a criterion in a group definition. Read/write  **PjGroupOn**.


## Syntax

 _expression_. **GroupOn**

 _expression_ A variable that represents an **GroupCriterion** object.


## Remarks

The  **GroupOn** property can be one of the **[PjGroupOn](pjgroupon-enumeration-project.md)** constants.


## Example

The following example adds a criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work that is complete. The GroupOn argument specifies that grouping is by a percentage interval.


```vb
Sub AddCriterionWithInterval() 
 ActiveProject.ResourceGroups("Response Pending").GroupCriteria.Add "% Work Complete", 
 True, CellColor:=pjRed, GroupOn:=pjGroupOnPctInterval, StartAt:=5, GroupInterval:=5 
End Sub
```


