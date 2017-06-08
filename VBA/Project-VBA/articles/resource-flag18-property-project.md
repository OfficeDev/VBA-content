---
title: Resource.Flag18 Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Flag18
ms.assetid: c8f1cf64-de8b-1b4c-30d7-6bf13b8ab5ea
ms.date: 06/08/2017
---


# Resource.Flag18 Property (Project)

 **True** if the flag associated with a **Resource** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag18**

 _expression_ A variable that represents a **Resource** object.


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


