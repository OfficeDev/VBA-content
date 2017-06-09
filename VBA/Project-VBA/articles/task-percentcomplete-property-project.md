---
title: Task.PercentComplete Property (Project)
ms.prod: project-server
api_name:
- Project.Task.PercentComplete
ms.assetid: fc698d7f-2dd9-9cbc-67ba-ff62e6db455c
ms.date: 06/08/2017
---


# Task.PercentComplete Property (Project)

Gets or sets the percent complete of a task. Read/write  **Variant**.


## Syntax

 _expression_. **PercentComplete**

 _expression_ A variable that represents a **Task** object.


## Example

The following example removes a resource from tasks that have two or more resources and are 85 percent complete.


```vb
Sub ReallocateResource() 
 
 Dim Entry As String ' The name of the resource to remove 
 Dim T As Task ' The task object used in For loop 
 Dim RA As Assignment ' The resource assignment object to the task 
 
 Entry = InputBox$("Enter a resource name:") 
 
 ' Remove the resource from 85 percent complete tasks with 2+ resources. 
 For Each T In ActiveProject.Tasks 
 If T.PercentComplete >= 85 And T.Resources.Count >= 2 Then 
 For Each RA In T.Assignments 
 If UCase(Entry) = UCase(RA.ResourceName) Then 
 RA.Delete 
 End If 
 Next 
 End If 
 Next T 
 
End Sub
```


