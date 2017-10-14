---
title: Application.WindowState Property (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowState
ms.assetid: 1a5d372d-9e05-80b4-6722-19781381d372
ms.date: 06/08/2017
---


# Application.WindowState Property (Project)

Gets or sets the state of the Project application window, where the state is maximized, minimized, or normal. Read/write  **PjWindowState**.


## Syntax

 _expression_. **WindowState**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **WindowState** property can be one of the **[PjWindowState](pjwindowstate-enumeration-project.md)** constants.

To change the state of a window within the application window, use the  **[WindowState](window-windowstate-property-project.md)** property of the **Window** object.


## Example

The following example minimizes the Project application window.


```vb
Sub MinimizeApplicationWindow() 
    Application.WindowState = pjMinimized 
End Sub
```


