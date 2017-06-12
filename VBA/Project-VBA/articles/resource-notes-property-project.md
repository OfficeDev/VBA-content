---
title: Resource.Notes Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Notes
ms.assetid: 63916a17-8ac0-e921-a29f-4d315c6cbc79
ms.date: 06/08/2017
---


# Resource.Notes Property (Project)

Gets or sets the notes for a resource. Read/write  **String**.


## Syntax

 _expression_. **Notes**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **Notes** property does not accept characters with an ASCII value less than 32, except for the carriage return (ASCII 13) and linefeed (ASCII 10) characters.


## Example

The following example adds a comment to the notes of the resource in the active cell.


 **Note**  If a resource is not selected, the code results in a run-time error 1004. 


```vb
Sub AddStatusNote() 
 Dim noStatus As String 
 noStatus = "No status report yet." 
 
 If ActiveCell.Resource.Notes = "" Then 
 ActiveCell.Resource.Notes = noStatus 
 Else 
 ActiveCell.Resource.Notes = ActiveCell.Resource.Notes _ 
 &; vbCrLf &; vbCrLf &; noStatus 
 End If 
 
End Sub
```


