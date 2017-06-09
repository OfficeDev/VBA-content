---
title: Task.PathSuccessor Property (Project)
ms.prod: project-server
ms.assetid: 827bf575-d93b-9959-c664-625c0e199699
ms.date: 06/08/2017
---


# Task.PathSuccessor Property (Project)
Gets a value that indicates whether the task is a successor of the selected task, when the  **Successors** item is selected in the **Task Path** drop-down list. Read-only **Boolean**.

## Syntax

 _expression_. **PathSuccessor**

 _expression_ A variable that represents a **Task** object.


## Remarks

The  **Task.PathSuccessor** property is related to the **Successors** item on the **Task Path** drop-down list, on the **FORMAT** tab, under **GANTT CHART TOOLS** on the ribbon. Task path is primarily a formatting feature in the Project client, where tasks in the Gantt chart have colors that depend on the current task selection and the relationship of a specified task to the selection. In Figure 1, the **Driving Predecessors** and **Driven Successors** items are selected in the **Task Path** drop-down list. When you select **T3**, the Gantt Chart shows that T1 is a driving predecessor task and T4 is a driven successor task.


**Figure 1. Using the task path properties to highlight tasks**

![Using the task path properties to highlight tasks](images/pj15_VBA_TaskPathDrivingPredecessor.gif)The  **PathSuccessor** property does not act like the **Successors** selection in the user interface. Instead, the **PathSuccessor** property is **True** whenboth of the following conditions are true: (a) the task is a successor of the selected task, and (b) the **Successors** item is selected in **Task Path**. You can manually select a task or use VBA to select a task, and then use VBA to check whether another task is a successor to the selected task. For example, if you select the third task as in Figure 1, and the  **Successors** item is selected in **Task Path**, the following statement prints  **True** in the **Immediate** window of the VBE.




```vb
? ActiveProject.Tasks(4).PathSuccessor
```

However, if the  **Successors** item is not selected, the previous statement prints **False**. Project does not have a VBA method that can set items in the  **Task Path** drop-down list.


## Example

The  **TestTaskPath** macro selects each task in a project, and then uses the four task path properties in turn to show how the other tasks relate to the selected task.


```vb
Option Explicit

Sub TestTaskPath()
    Dim t As Task
    Dim chkTsk As Task
    Dim selectedTaskId As Integer
    
    For Each t In ActiveProject.Tasks
        selectedTaskId = t.ID
        Application.SelectRow Row:=selectedTaskId, RowRelative:=False
            
        If Not (ActiveSelection.Tasks Is Nothing) Then
            Debug.Print
            
            With ActiveSelection.Tasks(1)
                Debug.Print "Selected task ID " &; .UniqueID &; ", name: " &; .Name
            End With
                        
            For Each chkTsk In ActiveProject.Tasks
                If Not (chkTsk.ID = selectedTaskId) Then
                    If chkTsk.PathPredecessor Then
                        Debug.Print vbTab &; chkTsk.Name &; ": PathPredecessor"
                    End If
                    If chkTsk.PathDrivingPredecessor Then
                        Debug.Print vbTab &; chkTsk.Name &; ": PathDrivingPredecessor"
                    End If
                    If chkTsk.PathSuccessor Then
                        Debug.Print vbTab &; chkTsk.Name &; ": PathSuccessor"
                    End If
                    If chkTsk.PathDrivenSuccessor Then
                        Debug.Print vbTab &; chkTsk.Name &; ": PathDrivenSuccessor"
                    End If
                End If
            Next chkTsk
        End If
    Next t
End Sub
```

For the project in Figure 1, if the  **Predecessors**,  **Driving Predecessors**,  **Successors**, and  **Driven Successors** items are all selected in **Task Path**, the  **TestTaskPath** macro has the following output:




```
Selected task ID 1, name: T1
    T2: PathSuccessor
    T2: PathDrivenSuccessor
    T3: PathSuccessor
    T3: PathDrivenSuccessor
    T4: PathSuccessor
    T4: PathDrivenSuccessor

Selected task ID 2, name: T2
    T1: PathPredecessor
    T1: PathDrivingPredecessor

Selected task ID 3, name: T3
    T1: PathPredecessor
    T1: PathDrivingPredecessor
    T4: PathSuccessor
    T4: PathDrivenSuccessor

Selected task ID 4, name: T4
    T1: PathPredecessor
    T1: PathDrivingPredecessor
    T3: PathPredecessor
    T3: PathDrivingPredecessor
```


## Property value

 **VARIANT**


## See also


#### Concepts


[Task Object](task-object-project.md)
#### Other resources


[PathDrivingPredecessor Property](task-pathdrivingpredecessor-property-project.md)
[PathPredecessor Property](task-pathpredecessor-property-project.md)
[PathDrivenSuccessor Property](task-pathdrivensuccessor-property-project.md)
