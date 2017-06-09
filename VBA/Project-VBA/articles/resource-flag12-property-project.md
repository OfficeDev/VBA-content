---
title: Resource.Flag12 Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Flag12
ms.assetid: 1ba606df-3606-b358-77b6-874551429cf4
ms.date: 06/08/2017
---


# Resource.Flag12 Property (Project)

 **True** if the flag associated with a **Resource** is set. Read/write **Variant**.


## Syntax

 _expression_. **Flag12**

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


