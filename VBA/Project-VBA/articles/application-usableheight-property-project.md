---
title: Application.UsableHeight Property (Project)
ms.prod: project-server
api_name:
- Project.Application.UsableHeight
ms.assetid: f0cd8b86-a619-022a-5e26-8d4c5e815af3
ms.date: 06/08/2017
---


# Application.UsableHeight Property (Project)

Gets the maximum height available for a project window in points. Read-only  **Double**.


## Syntax

 _expression_. **UsableHeight**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **UsableHeight** property equals the total amount of vertical space inside the main window minus the space taken up by the Ribbon, status bars, scroll bars, and the title bar.


## Example

The following example moves the windows of every open project inside the main window.


```vb
Sub FitWindows() 
 
 Dim W As Window ' The Window object used in For Each loop 
 
 For Each W In Application.Windows 
 ' Adjust the height of each window, if necessary. 
 If W.Height > UsableHeight Then 
 W.Height = UsableHeight 
 W.Top = 0 
 ' Adjust the vertical position of each window, if necessary. 
 ElseIf W.Top + W.Height > UsableHeight Then 
 W.Top = UsableHeight - W.Height 
 End If 
 
 ' Adjust the width of each window, if necessary. 
 If W.Width > UsableWidth Then 
 W.Width = UsableWidth 
 W.Left = 0 
 ' Adjust the horizontal position of each window, if necessary. 
 ElseIf W.Left + W.Width > UsableWidth Then 
 W.Left = UsableWidth - W.Width 
 End If 
 Next W 
 
End Sub
```


