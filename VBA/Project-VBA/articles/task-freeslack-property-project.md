---
title: Task.FreeSlack Property (Project)
keywords: vbapj.chm132289
f1_keywords:
- vbapj.chm132289
ms.prod: project-server
api_name:
- Project.Task.FreeSlack
ms.assetid: 714f6c83-bb4c-4a29-d9ea-e3f259d40c77
ms.date: 06/08/2017
---


# Task.FreeSlack Property (Project)

Gets the free slack for a task in minutes. Read-only  **Variant**.


## Syntax

 _expression_. **FreeSlack**

 _expression_ A variable that represents a **Task** object.


## Example

The following example eliminates free slack in the active project by changing the start dates of tasks with free slack.


```vb
Sub EliminateFreeSlack() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 For Each T In ActiveProject.Tasks 
 If T.FreeSlack > 0 Then 
 T.Start = Application.DateAdd(T.Start, T.FreeSlack) 
 End If 
 Next T 
 
End Sub
```


