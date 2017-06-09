---
title: Window.Close Method (Project)
ms.prod: project-server
api_name:
- Project.Window.Close
ms.assetid: 820f202b-d609-02e6-eff4-3368b9f93dd5
ms.date: 06/08/2017
---


# Window.Close Method (Project)

Closes a pane or window.


## Syntax

 _expression_. **Close**

 _expression_ A variable that represents a **Window** object.


## Example

The following example closes the lower pane of every open window.


```vb
Sub CloseWindowsOfActiveProject() 
 
 Dim W As Window 
 
 For Each W in Application.Windows 
 If Not (W.BottomPane Is Nothing) Then 
 W.BottomPane.Close 
 End If 
 Next W 
 
End Sub
```


