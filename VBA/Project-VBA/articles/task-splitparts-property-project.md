---
title: Task.SplitParts Property (Project)
ms.prod: project-server
api_name:
- Project.Task.SplitParts
ms.assetid: e4c62dce-4ee0-aff3-3248-f6b5b04b0c2d
ms.date: 06/08/2017
---


# Task.SplitParts Property (Project)

Gets a  **[SplitParts](splitpart-object-project.md)** collection that represents the portions of a split task. Read-only **SplitParts**.


## Syntax

 _expression_. **SplitParts**

 _expression_ A variable that represents a **Task** object.


## Example

The following example returns the number of parts for each task in the active project.


```vb
Sub CountTaskPortions() 
 Dim T As Task, HowMany As Long 
 
 For Each T In ActiveProject.Tasks 
 HowMany = 0 
 If Not (T Is Nothing) Then 
 HowMany = HowMany + T.SplitParts.Count 
 MsgBox T.Name &; ": " &; HowMany &; " task portion(s)" 
 End If 
 
 Next T 
 
End Sub
```


