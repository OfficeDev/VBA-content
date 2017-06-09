---
title: Resource.Level Method (Project)
ms.prod: project-server
api_name:
- Project.Resource.Level
ms.assetid: b6c7f694-0854-2ec0-48ec-91721cef993c
ms.date: 06/08/2017
---


# Resource.Level Method (Project)

Levels the resource.


## Syntax

 _expression_. **Level**

 _expression_ A variable that represents a **Resource** object.


## Example

The following example levels the resources of the selected tasks.


```vb
Sub LevelResourcesInSelectedTasks() 
    Dim T As Task ' Task object used in For Each loop 
    Dim A As Assignment ' Assignment object used in For Each loop 
 
    For Each T In ActiveSelection.Tasks 
        For Each A In T.Assignments 
            If ActiveProject.Resources(A.ResourceID).Overallocated Then 
                ActiveProject.Resources(A.ResourceID).Level 
            End If 
        Next A 
    Next T 
End Sub
```


