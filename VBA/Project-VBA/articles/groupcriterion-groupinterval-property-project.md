---
title: GroupCriterion.GroupInterval Property (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion.GroupInterval
ms.assetid: 1944776c-0150-d901-79f1-cfb7c0c698f7
ms.date: 06/08/2017
---


# GroupCriterion.GroupInterval Property (Project)

Gets or sets the interval for a field used as a criterion in a group definition. Read/write  **Variant**.


## Syntax

 _expression_. **GroupInterval**

 _expression_ A variable that represents an **GroupCriterion** object.


## Example

The following example adds a criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work that is complete. The interval for the group criterion is 5%.


```vb
Sub AddCriterionWithInterval() 
 ActiveProject.ResourceGroups("Response Pending").GroupCriteria.Add "% Work Complete", 
 True, CellColor:=pjRed, GroupOn:=pjGroupOnPctInterval, StartAt:=5, GroupInterval:=5 
End Sub
```


