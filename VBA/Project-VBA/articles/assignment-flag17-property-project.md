---
title: Assignment.Flag17 Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Flag17
ms.assetid: cda8dbba-c35c-86a8-348b-ed0ac4a15db5
ms.date: 06/08/2017
---


# Assignment.Flag17 Property (Project)

 **True** if the flag associated with an **Assignment** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag17**

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


