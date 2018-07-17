---
title: Window.Height Property (Project)
ms.prod: project-server
api_name:
- Project.Window.Height
ms.assetid: 4ed45f1f-c325-8a51-333c-28160d6b5f26
ms.date: 06/08/2017
---


# Window.Height Property (Project)

Gets or sets the height of a project window in points. Read/write  **Long**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **Window** object.


## Remarks

A window changes its height by moving its bottom edge, leaving the top edge unaffected.


## Example

The following example places the main window in the lower half of the screen.


```vb
Sub PlaceProjectInLowerScreenHalf() 
 
 Dim WindowWidth As Double 
 
 Application.WindowState = pjMaximized 
 WindowWidth = Application.Width 'Remember the width when maximized. 
 
 Application.Height = Application.Height / 2 
 Application.Top = Application.Height 
 
 'Ensure that the window uses all the available width. 
 If Application.Width < WindowWidth Then 
 Application.Width = WindowWidth 
 Application.Left = 0 
 End If 
 
End Sub
```


