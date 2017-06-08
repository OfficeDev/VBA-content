---
title: ActualStartDrivers.TotalDetectedCount Property (Project)
ms.prod: project-server
api_name:
- Project.ActualStartDrivers.TotalDetectedCount
ms.assetid: 188d79e3-3a1b-a0ed-e11b-3998334d6a17
ms.date: 06/08/2017
---


# ActualStartDrivers.TotalDetectedCount Property (Project)

Gets the total number of actual start drivers that affect the start date of a task. Read-only  **Long**.


## Syntax

 _expression_. **TotalDetectedCount**

 _expression_ A variable that represents an **ActualStartDrivers** object.


## Remarks

Actual start drivers are assignments that affect the start date of a task because they have actual work completed on the first day of the task.


## Example

The following example displays  **TotalDetectedCount** for each task in the active project. The example assumes there are no more than five assignments whose start dates are the same as the task's start date.


```vb
Sub b() 

 Dim T As Task 

 Dim count As Integer 

 For Each T In ActiveProject.Tasks 

 If T.RecalcFlags = 1 Then 

 MsgBox (T.StartDriver.ActualStartDrivers.TotalDetectedCount) 

 End If 

 Next T 

End Sub
```


## See also


#### Concepts


[ActualStartDrivers Collection Object](actualstartdrivers-object-project.md)

