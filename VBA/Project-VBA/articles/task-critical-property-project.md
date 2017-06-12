---
title: Task.Critical Property (Project)
keywords: vbapj.chm131694
f1_keywords:
- vbapj.chm131694
ms.prod: project-server
api_name:
- Project.Task.Critical
ms.assetid: 2282f751-adb3-d891-8d93-7e55723e2e7d
ms.date: 06/08/2017
---


# Task.Critical Property (Project)

 **True** if the task is on the critical path. Read-only **Boolean**.


## Syntax

 _expression_. **Critical**

 _expression_ A variable that represents a **Task** object.


## Example

The following example sets the highest priority for critical tasks in the active project.


```vb
Sub MakeCriticalTasksHighestPriority() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 For Each T In ActiveProject.Tasks 
 If T.Critical Then T.Priority = pjPriorityHighest 
 Next T 
 
End Sub
```


