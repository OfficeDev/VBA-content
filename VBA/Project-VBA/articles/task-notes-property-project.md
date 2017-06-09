---
title: Task.Notes Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Notes
ms.assetid: 65eecb2e-9116-2b00-8fb1-6df471a88f1d
ms.date: 06/08/2017
---


# Task.Notes Property (Project)

Gets or sets the notes for a task. Read/write  **String**.


## Syntax

 _expression_. **Notes**

 _expression_ A variable that represents a **Task** object.


## Remarks

The  **Notes** property does not accept characters with an ASCII value less than 32, except for the carriage return (ASCII 13) and linefeed (ASCII 10) characters.


## Example

The following example adds a comment to the notes of the task in the active cell.


 **Note**  If a task is not selected, the code results in a run-time error 1004. 


```vb
Sub AddDelayNote() 
 ActiveCell.Task.Notes = ActiveCell.Task.Notes &; vbCrLf &; vbCrLf &; "This task can be delayed." 
End Sub
```


