---
title: Project.ResourceViewList Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ResourceViewList
ms.assetid: d0acf85f-8a07-714d-614f-a18645177f40
ms.date: 06/08/2017
---


# Project.ResourceViewList Property (Project)

Gets a  **[List](list-object-project.md)** object representing all resource views in the active project. Read-only **List**.


## Syntax

 _expression_. **ResourceViewList**

 _expression_ A variable that represents a **Project** object.


## Example

The following example lists all the resource views in the active project.


```vb
Sub SeeAllResViews() 
 
 Dim Temp As Variant 
 Dim ResViewNames As String 
 
 For Each Temp In ActiveProject.ResourceViewList 
 ResViewNames = ResViewNames &; vbCrLf &; Temp 
 Next Temp 
 
 MsgBox ResViewNames 
 
End Sub
```


