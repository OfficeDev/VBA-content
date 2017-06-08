---
title: Assignment.Flag16 Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Flag16
ms.assetid: fc4034ce-15b2-42fa-a292-453f5b2abacd
ms.date: 06/08/2017
---


# Assignment.Flag16 Property (Project)

 **True** if the flag associated with an **Assignment** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag16**

 _expression_ A variable that represents an **Assignment** object.


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


