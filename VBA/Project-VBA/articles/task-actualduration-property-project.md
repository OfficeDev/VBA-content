---
title: Task.ActualDuration Property (Project)
keywords: vbapj.chm131381
f1_keywords:
- vbapj.chm131381
ms.prod: project-server
api_name:
- Project.Task.ActualDuration
ms.assetid: c0f56a31-acc1-215c-0737-d7ad755e0a96
ms.date: 06/08/2017
---


# Task.ActualDuration Property (Project)

Gets or sets the actual duration (in minutes) of a task. Read-only for summary tasks. Read/write  **Variant**.


## Syntax

 _expression_. **ActualDuration**

 _expression_ A variable that represents a **Task** object.


## Example

The following example marks the tasks in the active project that have actual durations that exceed a specified number of minutes.


```vb
Sub MarkLongDurationTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim Minutes As Long ' Duration entered by user 
 
 ' Prompt user for the actual duration, in minutes. 
 Minutes = Val(InputBox$("Enter the actual duration, in minutes: ")) 
 
 ' Don't do anything if the InputBox$ was cancelled. 
 If Minutes = 0 Then Exit Sub 
 
 ' Cycle through the tasks of the active project. 
 For Each T In ActiveProject.Tasks 
 ' Mark a task if it exceeds the duration. 
 If T.ActualDuration > Minutes Then T.Marked = True 
 Next T 
 
End Sub
```


