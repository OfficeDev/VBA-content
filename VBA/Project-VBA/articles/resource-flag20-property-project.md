---
title: Resource.Flag20 Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Flag20
ms.assetid: 3dbd0ffc-db53-fb14-e396-9f80c40fa5cf
ms.date: 06/08/2017
---


# Resource.Flag20 Property (Project)

 **True** if the flag associated with a **Resource** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag20**

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


