---
title: Application.UsableWidth Property (Project)
keywords: vbapj.chm132778
f1_keywords:
- vbapj.chm132778
ms.prod: project-server
api_name:
- Project.Application.UsableWidth
ms.assetid: ccc312da-6794-657d-7c76-e3e8549e2da7
ms.date: 06/08/2017
---


# Application.UsableWidth Property (Project)

Gets the maximum width available for a project window in points. Read-only Double.


## Syntax

 _expression_. **UsableWidth**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **UsableWidth** property equals the total amount of horizontal space inside the main window minus the space taken up by scroll bars.


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


