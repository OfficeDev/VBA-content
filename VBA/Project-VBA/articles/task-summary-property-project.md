---
title: Task.Summary Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Summary
ms.assetid: 252591e8-da5c-1b5e-a195-09deb44754af
ms.date: 06/08/2017
---


# Task.Summary Property (Project)

 **True** if the task is a summary task. Read-only **Boolean**.


## Syntax

 _expression_. **Summary**

 _expression_ A variable that represents a **Task** object.


## Example

The following example checks whether summary tasks in the active project have assignments. 


 **Note**  Assignments should not be made on summary tasks.


```vb
Sub CheckAssignmentsOnSummaryTasks() 
 Dim tsk As Task 
 Dim message As String 
 Dim numAssignments As Integer 
 Dim numSummaryTasksWithAssignments As Integer 
 Dim msgStyle As VbMsgBoxStyle 
 
 message = "" 
 numSummaryTasksWithAssignments = 0 
 
 For Each tsk In ActiveProject.Tasks 
 If tsk.Summary Then 
 numAssignments = tsk.Assignments.Count 
 If numAssignments > 0 Then 
 message = message &; "Summary task ID (" &; tsk.ID &; "): " &; tsk.Name _ 
 &; ": " &; numAssignments &; " assignments" &; vbCrLf 
 numSummaryTasksWithAssignments = numSummaryTasksWithAssignments + 1 
 End If 
 End If 
 Next tsk 
 
 If numSummaryTasksWithAssignments > 0 Then 
 message = "There are " &; numSummaryTasksWithAssignments _ 
 &; " summary tasks that have assignments." &; vbCrLf &; vbCrLf &; message 
 msgStyle = vbExclamation 
 Else 
 message = "No summary tasks have assignments." 
 msgStyle = vbInformation 
 End If 
 
 MsgBox message, msgStyle, "Summary Task Check" 
End Sub
```


