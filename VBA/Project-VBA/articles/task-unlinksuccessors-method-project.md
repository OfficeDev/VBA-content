---
title: Task.UnlinkSuccessors Method (Project)
ms.prod: project-server
api_name:
- Project.Task.UnlinkSuccessors
ms.assetid: ad3148f3-604c-aea9-f592-1f76372dffee
ms.date: 06/08/2017
---


# Task.UnlinkSuccessors Method (Project)

Removes one or more successors from the task.


## Syntax

 _expression_. **UnlinkSuccessors**( ** _Tasks_** )

 _expression_ A variable that represents a **Task** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Tasks_|Required|**Object**|Can be a **Task** or **Tasks** object, which specifies one or more tasks that are removed as successors.|

### Return Value

 **Nothing**


## Example

The following example removes the specified successor from every task in the active project.


```vb
Sub RemoveSuccessor() 
    Dim Entry As String  ' Successor specified by user 
    Dim SuccTask As Task ' Successor task object 
    Dim T As Task        ' Task object used in For Each loop 
    Dim S As Task        ' Successor (task object) used in loop 
 
    Entry = InputBox$("Enter the name of a successor to unlink from every task in this project.") 
    Set SuccTask = Nothing 
 
    ' Look for the name of the successor in tasks of the active project. 
    For Each T In ActiveProject.Tasks 
        If T.Name = Entry Then 
            Set SuccTask = T 
            Exit For 
        End If 
    Next T 
 
    ' Remove the successor from every task in the active project. 
    If Not (SuccTask Is Nothing) Then 
        For Each T In ActiveProject.Tasks 
            For Each S In T.SuccessorTasks 
                If S.Name = Entry Then 
                    T.UnlinkSuccessors SuccTask 
                    Exit For 
                End If 
            Next S 
        Next T 
    End If 
End Sub
```


