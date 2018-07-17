---
title: Project.ReadOnly Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ReadOnly
ms.assetid: 9ec47083-afb5-b51d-96e3-c460b02f2012
ms.date: 06/08/2017
---


# Project.ReadOnly Property (Project)

 **True** if a project has read-only access. Read-only **Boolean**.


## Syntax

 _expression_. **ReadOnly**

 _expression_ A variable that represents a **Project** object.


## Example

The following example copies projects with read-only access into new files with read-write access.


```vb
Sub CopyReadOnlyFiles() 
 
 Dim P As Project ' Project object used in loop 
 Dim OldName As String ' Name of project 
 Dim Path As String ' File path to project 
 Dim NewName As String ' New name of project 
 
 ' Check each open project for read-only access. 
 For Each P In Application.Projects 
 If P.ReadOnly Then ' See if project has read-only access. 
 OldName = P.Name ' Store its name. 
 Path = P.Path ' Store its path. 
 ' Create a new name for the file and save it. 
 NewName = "New " &; Left(OldName, Len(OldName) - 4) &; ".MPP" 
 P.Activate 
 FileSaveAs Path &; PathSeparator &; NewName 
 End If 
 Next P 
 
End Sub
```


