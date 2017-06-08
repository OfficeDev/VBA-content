---
title: Task.TaskDependencies Property (Project)
ms.prod: project-server
api_name:
- Project.Task.TaskDependencies
ms.assetid: 9c02fe5f-cb9e-a10e-bf9a-66b7600f8c64
ms.date: 06/08/2017
---


# Task.TaskDependencies Property (Project)

Gets a  **[TaskDependencies](taskdependency-object-project.md)** collection of dependent (predecessor and successor) tasks. Read-only **TaskDependencies**.


## Syntax

 _expression_. **TaskDependencies**

 _expression_ A variable that represents a **Task** object.


## Remarks

Each  **TaskDependency** object in the **TaskDependencies** collection includes the link type and link lag information between the tasks.


## Example

The following example examines each predecessor for the specified task and displays a message for each predecessor task that has a priority higher than "Medium."


```vb
Sub FindHighPriPreds() 
 Dim TaskDep As TaskDependency 
 
 For Each TaskDep In ActiveProject.Tasks("Write Requirements Brief").TaskDependencies 
 If TaskDep.From.Priority > 500 Then 
 MsgBox "Task #" &; TaskDep.From.ID &; " (" &; TaskDep.From.Name &; ") " &; _ 
 "has a priority higher than medium." 
 End If 
 Next TaskDep 
End Sub
```


