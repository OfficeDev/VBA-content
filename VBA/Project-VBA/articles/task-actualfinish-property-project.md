---
title: Task.ActualFinish Property (Project)
ms.prod: project-server
api_name:
- Project.Task.ActualFinish
ms.assetid: 183ce863-c7e9-77a7-1f0d-1452596b1b23
ms.date: 06/08/2017
---


# Task.ActualFinish Property (Project)

Gets or sets the actual finish date of a task. Read-only for summary tasks. Read/write  **Variant**.


## Syntax

 _expression_. **ActualFinish**

 _expression_ A variable that represents a **Task** object.


## Example

The following example prompts the user to set the actual finish dates of tasks in the active project.


```vb
Sub SetActualFinishForTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim Entry As String ' User's entry 
 
 For Each T In ActiveProject.Tasks 
 ' Loop until user enters a date or clicks Cancel. 
 Do While 1 
 Entry = InputBox$("Enter the actual finish date for " &; _ 
 T.Name &; ":") 
 
 If IsDate(Entry) Or Entry = Empty Then 
 Exit Do 
 Else 
 MsgBox ("You didn't enter a date; try again.") 
 End If 
 Loop 
 
 'If user didn't click Cancel, set the task's actual finish date. 
 If Entry <> Empty Then 
 T.ActualFinish = Entry 
 End If 
 
 Next T 
 
End Sub
```


