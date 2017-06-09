---
title: Project.CurrentView Property (Project)
ms.prod: project-server
api_name:
- Project.Project.CurrentView
ms.assetid: 002fc584-511e-0554-65f0-65dfd6b3dccb
ms.date: 06/08/2017
---


# Project.CurrentView Property (Project)

Gets the name of the active view for a project. Read-only  **String**.


## Syntax

 _expression_. **CurrentView**

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


