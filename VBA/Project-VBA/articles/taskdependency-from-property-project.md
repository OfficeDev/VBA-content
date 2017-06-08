---
title: TaskDependency.From Property (Project)
keywords: vbapj.chm132290
f1_keywords:
- vbapj.chm132290
ms.prod: project-server
api_name:
- Project.TaskDependency.From
ms.assetid: 76127fff-e8c0-f5b4-da5b-510a5f2222fa
ms.date: 06/08/2017
---


# TaskDependency.From Property (Project)

Gets a  **[Task](task-object-project.md)** object that is the predecessor in a task dependency. Read-only **Task**.


## Syntax

 _expression_. **From**

 _expression_ A variable that represents a **TaskDependency** object.


## Example

In the following example, the  **From** property appears to get both a **Task** object and a **Long** value. However, because **UniqueID** is the default property of a **Task** object, the second assignment using the **From** property is equivalent to the statement, `taskId = ActiveProject.Tasks(2).TaskDependencies(i).From.UniqueID`.


```vb
Sub TestDependenciesFrom() 
 Dim tsk As Task 
 Dim numDependencies As Integer 
 Dim taskId As Long 
 Dim i As Integer 
 
 numDependencies = ActiveProject.Tasks(2).TaskDependencies.Count 
 
 For i = 1 To numDependencies 
 Set tsk = ActiveProject.Tasks(2).TaskDependencies(i).From 
 Debug.Print tsk.Name 
 
 taskId = ActiveProject.Tasks(2).TaskDependencies(i).From 
 Debug.Print taskId 
 Next i 
End Sub
```


