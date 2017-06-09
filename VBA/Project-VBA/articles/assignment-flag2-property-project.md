---
title: Assignment.Flag2 Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Flag2
ms.assetid: a1659a3c-e5a9-0409-217c-3cb0be5c0818
ms.date: 06/08/2017
---


# Assignment.Flag2 Property (Project)

 **True** if the flag associated with an **Assignment** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag2**

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


