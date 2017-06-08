---
title: Assignment.Flag4 Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Flag4
ms.assetid: 16af5669-ced4-3f4b-063a-0755fcefbeb7
ms.date: 06/08/2017
---


# Assignment.Flag4 Property (Project)

 **True** if the flag associated with an **Assignment** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag4**

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


