---
title: Resource.AccrueAt Property (Project)
keywords: vbapj.chm131366
f1_keywords:
- vbapj.chm131366
ms.prod: project-server
api_name:
- Project.Resource.AccrueAt
ms.assetid: 760e1f6f-04b9-39e0-61a9-43af3813c473
ms.date: 06/08/2017
---


# Resource.AccrueAt Property (Project)

Gets or sets the way a task accrues the cost of a resource assigned to it. Read/write  **Variant**.


## Syntax

 _expression_. **AccrueAt**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **AccrueAt** property can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.


## Example

The following example sets the  **AccrueAt** property to **pjProrated** for each resource in the active project.


```vb
Sub SetProratedAccrueAt() 
 
 Dim R As Resource ' Resource object used in For Each loop 
 
 ' Cause tasks to accrue the cost of resources during the task. 
 For Each R In ActiveProject.Resources 
 R.AccrueAt = pjProrated 
 Next R 
 
End Sub
```


