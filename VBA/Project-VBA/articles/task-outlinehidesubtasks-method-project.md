---
title: Task.OutlineHideSubTasks Method (Project)
ms.prod: project-server
api_name:
- Project.Task.OutlineHideSubTasks
ms.assetid: 877e8248-3e3f-1816-0799-52fb5cda1d60
ms.date: 06/08/2017
---


# Task.OutlineHideSubTasks Method (Project)

Hides the subtasks of the selected task or tasks.


## Syntax

 _expression_. **OutlineHideSubTasks**

 _expression_ A variable that represents a **Task** object.


## Example

The following example collapses the entire outline of the first task.


```vb
Sub OutlineHideAllSubtasks() 
 ActiveProject.Tasks(1).OutlineHideSubtasks 
End Sub
```


