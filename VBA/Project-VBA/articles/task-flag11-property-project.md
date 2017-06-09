---
title: Task.Flag11 Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Flag11
ms.assetid: defeee71-a899-81b5-4107-602fd0e9373c
ms.date: 06/08/2017
---


# Task.Flag11 Property (Project)

Gets or sets the value of a task flag custom field. Read/write  **Variant**.


## Syntax

 _expression_. **Flag11**

 _expression_ A variable that represents a **Task** object.


## Example

The following example deletes all the tasks that have the  **Flag1** set to **True**.


```vb
Sub DeleteNonEssentialTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 ' Delete nonessential tasks in the active project. 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.Flag1 = True Then T.Delete 
 End If 
 Next T 
 
End Sub
```


