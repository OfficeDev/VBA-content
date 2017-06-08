---
title: Project.ResourceTableList Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ResourceTableList
ms.assetid: 3d6c7995-4527-1597-ec56-c75d59be131a
ms.date: 06/08/2017
---


# Project.ResourceTableList Property (Project)

Gets a  **[List](list-object-project.md)** object representing all resource tables in the project. Read-only **List**.


## Syntax

 _expression_. **ResourceTableList**

 _expression_ A variable that represents a **Project** object.


## Example

The following example lists all the resource tables in the active project.


```vb
Sub SeeAllResTables() 
 
 Dim Temp As Variant 
 Dim ResTableNames As String 
 
 For Each Temp In ActiveProject.ResourceTableList 
 ResTableNames = ResTableNames &; vbCrLf &; Temp 
 Next Temp 
 
 MsgBox ResTableNames 
 
End Sub
```


