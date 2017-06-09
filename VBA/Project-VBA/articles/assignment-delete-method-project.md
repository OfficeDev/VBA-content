---
title: Assignment.Delete Method (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Delete
ms.assetid: 3147c0e0-239c-75d2-cae9-c299412190e2
ms.date: 06/08/2017
---


# Assignment.Delete Method (Project)

Deletes the  **Assignment** object from an **Assignments** collection.


## Syntax

 _expression_. **Delete**

 _expression_ A variable that represents an **Assignment** object.


## Example

The following example deletes every resource assignment in the active project.


```vb
Sub DeleteAssignments() 
 
 Dim RA As Assignment ' Assignment object for resources 
 Dim T As Task ' Task object 
 
 ' Delete resource assignments. 
 For Each T in ActiveProject.Tasks 
 For Each RA in T.Assignments 
 RA.Delete 
 Next RA 
 Next T 
 
End Sub
```


