---
title: Assignment.StartVariance Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.StartVariance
ms.assetid: 080f4dea-76aa-5438-e44a-ab71732b30b1
ms.date: 06/08/2017
---


# Assignment.StartVariance Property (Project)

Gets the variance (in minutes) between the baseline start date and the start date of the assignment. Read-only  **Variant**.


## Syntax

 _expression_. **StartVariance**

 _expression_ A variable that represents an **Assignment** object.


## Example

The following example displays the number of tasks in the active project that have started late.


```vb
Sub CountLateAssignments() 
 
 Dim a As Assignment 
 Dim t As Task 
 Dim numLateAssignments As Long 
 Dim lateAssignments As String 
 Dim daysLate As Single 
 
 numLateAssignments = 0 
 
 ' Look for late tasks in the active project. 
 For Each t In ActiveProject.Tasks 
 For Each a In t.Assignments 
 If a.BaselineStart < ActiveProject.CurrentDate And a.StartVariance > 0 Then 
 numLateAssignments = numLateAssignments + 1 
 daysLate = Round(a.StartVariance / 1440, 1) 
 lateAssignments = lateAssignments &; vbCrLf &; vbTab &; t.Name _ 
 &; ": resource " &; a.Resource.Name &; ": " &; daysLate &; " days" 
 End If 
 Next a 
 Next t 
 
 MsgBox "There are " &; numLateAssignments &; " late assignments in this project: " &; lateAssignments 
 
End Sub
```


