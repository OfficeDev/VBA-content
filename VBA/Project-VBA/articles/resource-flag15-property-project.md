---
title: Resource.Flag15 Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Flag15
ms.assetid: 5adc1e9e-cf5e-d19a-8e12-b80110766b88
ms.date: 06/08/2017
---


# Resource.Flag15 Property (Project)

 **True** if the flag associated with a **Resource** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag15**

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


