---
title: TaskDependency.LagType Property (Project)
ms.prod: project-server
api_name:
- Project.TaskDependency.LagType
ms.assetid: 0c055a94-ea5f-1267-0b61-d3a50c6bc9b4
ms.date: 06/08/2017
---


# TaskDependency.LagType Property (Project)

Gets the unit of lag time between linked tasks. Read-only  **[PjFormatUnit](pjformatunit-enumeration-project.md)**.


## Syntax

 _expression_. **LagType**

 _expression_ A variable that represents a **TaskDependency** object.


## Remarks

In the  **Lag** property, string values default to days unless otherwise specified. Non-string values are interpreted as minutes. To specify lead time between tasks, use a negative value for the **Lag** property.


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


