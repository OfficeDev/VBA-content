---
title: Task.Name Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Name
ms.assetid: 2df034b0-13bc-f912-abbc-6b97b8c8d5ed
ms.date: 06/08/2017
---


# Task.Name Property (Project)

Gets or sets the name of a  **Task** object. Read/write **String**.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Task** object.


## Example

The following example displays the task names that contain the specified text.


```vb
Sub NameExample() 
    Dim t As Task 
    Dim x As String 
    Dim y As String 
 
    x = InputBox$("Search for tasks that include the following text in their names:") 
 
    If Not x = "" Then 
        For Each t In ActiveProject.Tasks 
            If InStr(1, t.Name, x, 1) Then 
                y = y &; vbCrLf &; t.ID &; ": " &; t.Name 
            End If 
        Next t 
 
        If Len(y) = 0 Then 
            MsgBox "No tasks with the text " &; x &; " found in the project", vbExclamation 
        Else 
            MsgBox y 
        End If 
    End If 
End Sub
```


