---
title: Task.Rollup Property (Project)
keywords: vbapj.chm132588
f1_keywords:
- vbapj.chm132588
ms.prod: project-server
api_name:
- Project.Task.Rollup
ms.assetid: 8f29afc1-85ec-d835-bc08-7311e9063ae4
ms.date: 06/08/2017
---


# Task.Rollup Property (Project)

 **True** if the dates of a subtask appear on its corresponding summary task bar. Read/write **Variant**.


## Syntax

 _expression_. **Rollup**

 _expression_ A variable that represents a **Task** object.


## Remarks

The  **Rollup** property must be **True** on the summary task as well as the subtasks for the rollup to occur.


## Example

The following example sets the  **Rollup** property to **True** for milestone tasks, and to **False** for other tasks in the active project.


```vb
Sub DisplayMilestonesInSummaryBars() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 ' Cycle through tasks in active project. 
 For Each T In ActiveProject.Tasks 
 ' If task is a milestone or a summary, set its Rollup property to True. 
 If T.Summary Or T.Milestone Then 
 T.Rollup = True 
 ' If task isn't a summary task or milestone, set its Rollup property to False. 
 Else 
 T.Rollup = False 
 End If 
 Next T 
 
End Sub
```


