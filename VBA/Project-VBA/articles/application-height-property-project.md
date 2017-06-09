---
title: Application.Height Property (Project)
ms.prod: project-server
api_name:
- Project.Application.Height
ms.assetid: e980a85d-218c-b82d-1043-9670cab23560
ms.date: 06/08/2017
---


# Application.Height Property (Project)

Gets or sets the height of the main window in points. Read/write  **Long**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents an **Application** object.


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


