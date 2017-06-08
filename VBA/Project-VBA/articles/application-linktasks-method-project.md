---
title: Application.LinkTasks Method (Project)
keywords: vbapj.chm210
f1_keywords:
- vbapj.chm210
ms.prod: project-server
api_name:
- Project.Application.LinkTasks
ms.assetid: cc41c963-533c-97bf-8301-388bb2aaf746
ms.date: 06/08/2017
---


# Application.LinkTasks Method (Project)

Links the selected tasks in the Gantt Chart, Calendar, Task Sheet, or Task Usage view.


## Syntax

 _expression_. **LinkTasks**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Example

The following example a couple of tasks and links them.


```vb
Sub Link_Tasks() 
 
 'Activate Task Sheet view 
 ViewApply Name:="Task Sheet" 
 
 ' Create a coupe of tasks 
 RowInsert 
 SetTaskField Field:="Name", Value:="TestTask-2" 
 SetTaskField Field:="Duration", Value:="5" 
 
 RowInsert 
 SetTaskField Field:="Name", Value:="TestTask-1" 
 SetTaskField Field:="Duration", Value:="10" 
 
 'Select tasks 
 SelectRow 
 SelectRow Row:=1, Add:=True 
 
 'Link the two tasks 
 LinkTasks 
 
 'delete the tasks 
 ActiveProject.Tasks("TestTask-1").Delete 
 ActiveProject.Tasks("TestTask-2").Delete 
End Sub
```


