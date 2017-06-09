---
title: Selection.Tasks Property (Project)
ms.prod: project-server
api_name:
- Project.Selection.Tasks
ms.assetid: 8f58ea8e-a3a1-f5aa-ad5d-6447fe777453
ms.date: 06/08/2017
---


# Selection.Tasks Property (Project)

Gets a  **[Tasks](task-object-project.md)** collection representing the tasks in the selection. Read-only **Tasks**.


## Syntax

 _expression_. **Tasks**

 _expression_ A variable that represents a **Selection** object.


## Example

The following example displays the name of every task in the selection.


```vb
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveSelection.Tasks 
 Names = Names &; T.Name &; vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```


