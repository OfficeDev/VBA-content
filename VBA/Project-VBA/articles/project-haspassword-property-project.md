---
title: Project.HasPassword Property (Project)
ms.prod: project-server
api_name:
- Project.Project.HasPassword
ms.assetid: 2c00e008-94d9-5d0a-d3b9-dcb57af04a19
ms.date: 06/08/2017
---


# Project.HasPassword Property (Project)

 **True** if a project has a password. Read-only **Boolean**.


## Syntax

 _expression_. **HasPassword**

 _expression_ A variable that represents a **Project** object.


## Remarks




 **Note**  Project can open project files stored in an ODBC database that have a password, but cannot save files to a database. 


## Example

The following example displays a list of open projects that have passwords.


```vb
Sub ListProjectsWithPasswords() 
    Dim P As Project ' Project object used in For Each loop 
    Dim NameList As String ' Names of projects with passwords 
 
    ' Check each open project for passwords. 
    For Each P in Application.Projects 
        ' If a project has a password, add its name to the list. 
        If P.HasPassword Then 
            NameList = NameList &; P.Name &; vbCrLf 
        End If 
    Next P 
 
    ' Display information about projects with passwords. 
    If NameList = "" Then 
        MsgBox("No open projects have passwords.") 
    Else 
        MsgBox("The following open projects have passwords: " &; vbCrLf &; vbCrLf &; NameList) 
    End If 
End Sub
```


