---
title: Task.Milestone Property (Project)
keywords: vbapj.chm132409
f1_keywords:
- vbapj.chm132409
ms.prod: project-server
api_name:
- Project.Task.Milestone
ms.assetid: 246b3d92-43d7-850b-ab7c-8c314ca42aa9
ms.date: 06/08/2017
---


# Task.Milestone Property (Project)

 **True** if the task is a milestone. Read/write **Variant**.


## Syntax

 _expression_. **Milestone**

 _expression_ A variable that represents a **Task** object.


## Example

The following example marks as milestones any tasks in the active project with names that begin with the word "Inspection."


```vb
Sub MarkInspectionTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim MilestoneName As String 
 Dim NameLength As Integer 
 
 MilestoneName = "Inspection" 
 NameLength = Len(MilestoneName) 
 
 For Each T In ActiveProject.Tasks 
 ' If the task's name begins with Inspection, it's a milestone. 
 If UCase(Left(T.Name, NameLength)) = UCase(MilestoneName) Then 
 T.Milestone = True 
 End If 
 Next T 
 
End Sub
```


