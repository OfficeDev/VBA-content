---
title: Task.WBSPredecessors Property (Project)
keywords: vbapj.chm132764
f1_keywords:
- vbapj.chm132764
ms.prod: project-server
api_name:
- Project.Task.WBSPredecessors
ms.assetid: e4f71c96-44dc-9074-b424-2b4a7f939988
ms.date: 06/08/2017
---


# Task.WBSPredecessors Property (Project)

Gets the work breakdown structure (WBS) codes of the task predecessors, separated by the list separator. Read-only  **String**.


## Syntax

 _expression_. **WBSPredecessors**

 _expression_ A variable that represents a **Task** object.


## Example

The following example queries the user for a task ID and then provides a more user-friendly breakdown of its predecessors' WBS codes.


```vb
Sub EnumeratePredecessors() 
 Dim Task As Task 
 Dim PredTasks As Tasks 
 Dim ID As Long 
 Dim Predecessors As String 
 Dim List As String 
 Dim Count As Integer 
 
 ID = CLng(InputBox$("Enter the ID number of the task you wish to examine:")) 
 
 Set Task = ActiveProject.Tasks(ID) 
 Set PredTasks = Task.PredecessorTasks 
 Predecessors = Task.WBSPredecessors 
 Count = 1 
 
 If PredTasks.Count = 0 Then 
 List = "Task " &; Task.UniqueID &; ", " &; Task.Name &; ", has no predecessors." 
 Else 
 List = "Predecessors to task " &; Task.UniqueID &; ", " &; Task.Name &; ":" &; vbCrLf &; vbCrLf 
 Do While InStr(Predecessors, ListSeparator) <> 0 
 List = List &; PredTasks(Count).Name &; ": " &; Mid$(Predecessors, 1, InStr(Predecessors, ListSeparator) - 1) &; vbCrLf 
 Predecessors = Right$(Predecessors, Len(Predecessors) - InStr(Predecessors, ListSeparator)) 
 Count = Count + 1 
 Loop 
 List = List &; PredTasks(Count).Name &; ": " &; Predecessors 
 End If 
 
 MsgBox List 
 
 Set PredTasks = Nothing 
 Set Task = Nothing 
End Sub
```


