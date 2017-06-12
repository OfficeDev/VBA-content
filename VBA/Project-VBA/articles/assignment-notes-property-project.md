---
title: Assignment.Notes Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Notes
ms.assetid: 91915e62-bd93-3671-a232-05cb99836428
ms.date: 06/08/2017
---


# Assignment.Notes Property (Project)

Gets or sets the notes for an assignment. Read/write  **String**.


## Syntax

 _expression_. **Notes**

 _expression_ A variable that represents an **Assignment** object.


## Remarks

The  **Notes** property does not accept characters with an ASCII value less than 32, except for the carriage return (ASCII 13) and linefeed (ASCII 10) characters.


## Example

The following example adds a comment to the notes of the assignment in the active cell.


 **Note**  If an assignment is not selected, the code results in a run-time error 1004. 


```vb
Sub AddStatusNote() 
 Dim noStatus As String 
 noStatus = "No status report yet." 
 
 If ActiveCell.Assignment.Notes = "" Then 
 ActiveCell.Assignment.Notes = "No status report yet." 
 Else 
 ActiveCell.Assignment.Notes = ActiveCell.Assignment.Notes &; vbCrLf &; vbCrLf &; "No status report yet." 
 End If 
 
End Sub
```


