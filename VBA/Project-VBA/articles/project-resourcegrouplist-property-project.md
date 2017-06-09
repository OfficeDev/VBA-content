---
title: Project.ResourceGroupList Property (Project)
keywords: vbapj.chm132565
f1_keywords:
- vbapj.chm132565
ms.prod: project-server
api_name:
- Project.Project.ResourceGroupList
ms.assetid: a64fe8c8-e75f-3cab-e77a-54fc6c1bf0a5
ms.date: 06/08/2017
---


# Project.ResourceGroupList Property (Project)

Gets a  **[List](list-object-project.md)** object representing the resource groups in the active project. Read-only **List**.


## Syntax

 _expression_. **ResourceGroupList**

 _expression_ A variable that represents a **Project** object.


## Example

The following example lists all the resource filters in the active project.


```vb
Sub SeeAllResGroups() 
 
 Dim Temp As Variant 
 Dim ResGroupNames As String 
 
 For Each Temp In ActiveProject.ResourceGroupList 
 ResGroupNames = ResGroupNames &; vbCrLf &; Temp 
 Next Temp 
 
 MsgBox ResGroupNames 
 
End Sub
```


