---
title: TaskDependency.Lag Property (Project)
keywords: vbapj.chm132365
f1_keywords:
- vbapj.chm132365
ms.prod: project-server
api_name:
- Project.TaskDependency.Lag
ms.assetid: d3370ea3-5485-24d5-e363-ec4b5a0ec95b
ms.date: 06/08/2017
---


# TaskDependency.Lag Property (Project)

The duration of lag time between linked tasks. Read/write  **Variant**.


## Syntax

 _expression_. **Lag**

 _expression_ A variable that represents a **TaskDependency** object.


## Remarks

To specify lead time between tasks, use a negative value. String values default to days unless otherwise specified. Non-string values are interpreted as minutes.


## Example

To use the  **SetLagWeeks** macro, create two tasks, where Task 2 is linked to Task 1. When you run the macro, the **Immediate** window shows 4800 and 9, where the lag time is 4800 minutes and the type of lag is 9 ( **pjWeeks** ).


```vb
Sub SetLagWeeks() 
 Dim tsk As Task 
 Set tsk = ActiveProject.Tasks(2) 
 
 tsk.TaskDependencies(1).Lag = "2w" 
 
 Debug.Print tsk.TaskDependencies(1).Lag 
 Debug.Print tsk.TaskDependencies(1).LagType 
End Sub
```


