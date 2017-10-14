---
title: Task.Priority Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Priority
ms.assetid: 8680e903-a03e-cf83-70e7-fc856297dae2
ms.date: 06/08/2017
---


# Task.Priority Property (Project)

Gets or sets the priority for the task. Read/write  **Variant**.


## Syntax

 _expression_. **Priority**

 _expression_ A variable that represents a **Task** object.


## Remarks

The  **Priority** property can be a value from 0 to 1000. A value of 1000 has the effect that the task is not leveled in a leveling operation.Because **Priority** is a **Variant**, you can set the priority of task 2 to 900, for example, with the following code: `activeproject.Tasks(2).Priority = "Highest"`. The following table shows the string values and the corresponding integer values for the  **Priority** property.


 **Note**  Do not use the  **[PjPriority](pjpriority-enumeration-project.md)** constants, which have values only from 0 to 9 for some previous versions of Project.


|||
|:-----|:-----|
|**String**|**Priority value**|
|"Do not level"|1000|
|"Highest|900|
|"Very high"|800|
|"Higher"|700|
|"High"|600|
|"Medium"|500|
|"Low"|400|
|"Lower"|300|
|"Very low"|200|
|"Lowest"|100|
Project uses the  **Priority** property of the project summary task (task 0) to determine how to treat tasks when leveling resources across multiple projects. If two projects have equal priorities, the priority for individual tasks is used. You can set the project priority in the **Project Information** dialog box, or show the project summary task on the Gantt chart, select the task, and then use a statement such as `ActiveCell.Task.Priority = 700`.


## Example

The following example sets the tasks on the critical path to a very high priority in the active project.


```vb
Sub SetPriorityOfCriticalTasks() 
    Dim T As Task ' Task object used in For Each loop 
 
    ' Look for tasks on the critical path. 
    For Each T In ActiveProject.Tasks 
        If T.Critical = True Then 
            T.Priority = 800 
        End If 
    Next T 
End Sub
```


