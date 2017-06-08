---
title: Task.StartVariance Property (Project)
ms.prod: project-server
api_name:
- Project.Task.StartVariance
ms.assetid: 8ec7f5c9-62c4-36fd-d245-4a2bf21fd7bd
ms.date: 06/08/2017
---


# Task.StartVariance Property (Project)

Gets the variance (in minutes) between the baseline start date and the start date of the task. Read-only  **Variant**.


## Syntax

 _expression_. **StartVariance**

 _expression_ A variable that represents a **Task** object.


## Example

The following example displays the number of tasks and task names in the active project that have started late.


```vb
Sub CountLateTasks() 
 
 Dim t As Task 
 Dim numLateTasks As Long 
 Dim lateTasks As String 
 Dim daysLate As Single 
 
 numLateTasks = 0 
 
 ' Look for late tasks in the active project. 
 For Each t In ActiveProject.Tasks 
 If t.BaselineStart < ActiveProject.CurrentDate And t.StartVariance > 0 Then 
 numLateTasks = numLateTasks + 1 
 daysLate = Round(t.StartVariance / 1440, 1) 
 lateTasks = lateTasks &; vbCrLf &; vbTab &; t.Name _ 
 &; ": " &; daysLate &; " days" 
 End If 
 Next t 
 
 MsgBox "There are " &; numLateTasks &; " late tasks in this project: " &; lateTasks 
 
End Sub
```


