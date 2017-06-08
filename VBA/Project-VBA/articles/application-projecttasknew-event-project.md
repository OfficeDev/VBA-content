---
title: Application.ProjectTaskNew Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectTaskNew
ms.assetid: 40e9d8da-f863-a73e-56e9-bb89327142fb
ms.date: 06/08/2017
---


# Application.ProjectTaskNew Event (Project)

Occurs when a new task is created.


## Syntax

 _expression_. **ProjectTaskNew**( ** _pj_**, ** _ID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project where the task was created.|
| _ID_|Required|**Long**|The ID of the task that was created.|

### Return Value

nothing


## Remarks

You can trap project-level events from outside VBA.


## Example

 The following example shows how the **ProjectTaskNew** event can trap project-level events. In this case, the **App_ProjectTaskNew** event handler sets the global **ProjTaskNew** variable that the **Change** event handler uses. You can use similar code with the **[ProjectResourceNew](application-projectresourcenew-event-project.md)** and **[ProjectAssignmentNew](application-projectassignmentnew-event-project.md)** events.


1. Create a class module named  **EventClassModule**, and then insert the following code:
    
  ```
  Option Explicit 
Option Base 1 
 
Public WithEvents App As Application 
Public WithEvents Proj As Project 
 
Dim NewTaskIDs() As Integer 
Dim NumNewTasks As Integer 
 
Dim ProjTaskNew As Boolean 
 
Private Sub App_ProjectTaskNew(ByVal pj As Project, ByVal ID As Long) 
    NumNewTasks = NumNewTasks + 1 
 
    If ProjTaskNew Then 
        ReDim Preserve NewTaskIDs(NumNewTasks) As Integer 
    Else 
        ReDim NewTaskIDs(NumNewTasks) As Integer 
    End If 
 
    NewTaskIDs(NumNewTasks) = ID 
 
    ProjTaskNew = True 
End Sub 
 
Private Sub Proj_Change(ByVal pj As Project) 
    Dim NewTaskID As Variant 
 
    If ProjTaskNew Then 
        For Each NewTaskID In NewTaskIDs 
            MsgBox "New Task Name: " &; ActiveProject.Tasks.UniqueID(NewTaskID).Name 
        Next NewTaskID 
 
        NumNewTasks = 0 
 
        ProjTaskNew = False 
    End If 
End Sub
```


    
    
2. In a separate module, insert the following code:
    
  ```
  Option Explicit 
 
Dim X As New EventClassModule 
 
Sub Initialize_App() 
    Set X.App = MSProject.Application 
    Set X.Proj = Application.ActiveProject 
End Sub
  ```


    
    
3. Run the  **Initialize_App** procedure to start listening to the events.
    
4. Create a task. The event handler shows a message box every time a new task is added.
    



