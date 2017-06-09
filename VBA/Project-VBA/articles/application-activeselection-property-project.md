---
title: Application.ActiveSelection Property (Project)
keywords: vbapj.chm131378
f1_keywords:
- vbapj.chm131378
ms.prod: project-server
api_name:
- Project.Application.ActiveSelection
ms.assetid: aa72b337-4031-a970-0921-d1d60f66096e
ms.date: 06/08/2017
---


# Application.ActiveSelection Property (Project)

Gets a  **[Selection](selection-object-project.md)** object that represents the active selection. Read-only **Selection**.


## Syntax

 _expression_. **ActiveSelection**

 _expression_ A variable that represents an **Application** object.


## Example

The following example displays the name of each selected task in a message box. Running this example without a valid selection results in a trappable error (error code 424).


```vb
Sub SelectedTasks() 
 
 Dim T As Task 
 
 If Not (ActiveSelection.Tasks Is Nothing) Then 
 For Each T In ActiveSelection.Tasks 
 ' Test for blank task row 
 If Not (T Is Nothing) Then 
 MsgBox T.Name 
 End If 
 Next T 
 End If 
 
End Sub
```


