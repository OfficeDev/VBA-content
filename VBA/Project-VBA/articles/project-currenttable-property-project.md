---
title: Project.CurrentTable Property (Project)
ms.prod: project-server
api_name:
- Project.Project.CurrentTable
ms.assetid: 7b80d451-bf37-7b1c-62b4-7ee0e7fd0e63
ms.date: 06/08/2017
---


# Project.CurrentTable Property (Project)

Gets the name of the active table for a project. Read-only  **String**.


## Syntax

 _expression_. **CurrentTable**

 _expression_ A variable that represents a **Project** object.


## Example

The following example displays the names of the active view, table, and filter in a dialog box.


```vb
Sub ViewDetails() 
 
    Dim Temp As String 
     
    Temp = "View: " &; ActiveProject.CurrentView &; vbCrLf 
    Temp = Temp &; "Table:" &; ActiveProject.CurrentTable &; vbCrLf 
    Temp = Temp &; "Filter: " &; ActiveProject.CurrentFilter 
    MsgBox Temp 
End Sub
```


