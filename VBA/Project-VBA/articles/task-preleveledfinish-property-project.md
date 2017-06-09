---
title: Task.PreleveledFinish Property (Project)
ms.prod: project-server
api_name:
- Project.Task.PreleveledFinish
ms.assetid: edcb110a-41b7-c2ad-0382-d88cf5f3708c
ms.date: 06/08/2017
---


# Task.PreleveledFinish Property (Project)

Gets the finish date of a task before leveling occurred. Read-only  **Variant**.


## Syntax

 _expression_. **PreleveledFinish**

 _expression_ A variable that represents a **Task** object.


## Example

The following example calculates the difference, if any, between the projected finish date and the projected finish date before the task was leveled for each task in the project, and then displays those that changed.


```vb
Sub DateDifferences() 
 Dim T As Task, Results As String 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 ' Tasks that have never been leveled return "NA" 
 If T.PreleveledFinish <> "NA" And T.Finish <> T.PreleveledFinish Then 
 Results = Results &; T.Name &; ": " &; _ 
 DateDiff("d", T.PreleveledFinish, T.Finish) &; _ 
 " days" &; vbCrLf 
 End If 
 End If 
 Next T 
 
 If Results <> "" Then MsgBox Results 
 
End Sub
```


