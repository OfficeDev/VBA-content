---
title: Window.WindowState Property (Project)
ms.prod: project-server
api_name:
- Project.Window.WindowState
ms.assetid: b1c0616c-7377-356e-446d-ee2d2f490e15
ms.date: 06/08/2017
---


# Window.WindowState Property (Project)

Gets or sets the state the window, where the state is maximized or normal. Read/write  **PjWindowState**.


## Syntax

 _expression_. **WindowState**

 _expression_ A variable that represents a **Window** object.


## Remarks

The  **WindowState** property can be one of the following **[PjWindowState](pjwindowstate-enumeration-project.md)** constants: **pjMaximized** or **pjNormal**. The **pjMinimized** value has no effect on a window within the Project application.

To change the state of the application window, use the  **[WindowState](application-windowstate-property-project.md)** property of the **Application** object.


## Example

The following example maximizes the active window.


```vb
Sub MaximizeProjectWindow() 
 ActiveWindow.WindowState = pjMaximized 
End Sub
```


