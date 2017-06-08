---
title: Task.WBSSuccessors Property (Project)
keywords: vbapj.chm132816
f1_keywords:
- vbapj.chm132816
ms.prod: project-server
api_name:
- Project.Task.WBSSuccessors
ms.assetid: 4d435645-8437-af81-ad38-eca0c27cfd80
ms.date: 06/08/2017
---


# Task.WBSSuccessors Property (Project)

Gets the work breakdown structure (WBS) codes of the task successors, separated by the list separator. Read-only  **String**.


## Syntax

 _expression_. **WBSSuccessors**

 _expression_ A variable that represents a **Task** object.


## Example

The following example queries the user for a task ID and then provides a more user-friendly breakdown of its successors' WBS codes.


```vb
Sub EnumerateSuccessors() 
 Dim Task As Task 
 Dim SuccTasks As Tasks 
 Dim ID As Long 
 Dim Successors As String 
 Dim List As String 
 Dim Count As Integer 
 
 ID = CLng(InputBox$("Enter the ID number of the task you wish to examine:")) 
 
 Set Task = ActiveProject.Tasks(ID) 
 Set SuccTasks = Task.SuccessorTasks 
 Successors = Task.WBSSuccessors 
 Count = 1 
 
 If SuccTasks.Count = 0 Then 
 List = "Task " &; Task.UniqueID &; ", " &; Task.Name &; ", has no successors." 
 Else 
 List = "Successors to task " &; Task.UniqueID &; ", " &; Task.Name &; ":" &; vbCrLf &; vbCrLf 
 Do While InStr(Successors, ListSeparator) <> 0 
 List = List &; SuccTasks(Count).Name &; ": " &; Mid$(Successors, 1, InStr(Successors, ListSeparator) - 1) &; vbCrLf 
 Successors = Right$(Successors, Len(Successors) - InStr(Successors, ListSeparator)) 
 Count = Count + 1 
 Loop 
 List = List &; SuccTasks(Count).Name &; ": " &; Successors 
 End If 
 
 MsgBox List 
 
 Set SuccTasks = Nothing 
 Set Task = Nothing 
End Sub
```


