---
title: Window.Top Property (Project)
ms.prod: project-server
api_name:
- Project.Window.Top
ms.assetid: 60aca1d3-5ca5-093f-7828-39974300257f
ms.date: 06/08/2017
---


# Window.Top Property (Project)

Gets or sets the distance in points of the window below the top edge of the window display area. Read/write  **Long**.


## Syntax

 _expression_. **Top**

 _expression_ A variable that represents a **Window** object.


## Remarks

The window display area is below the ribbon in Project and Project. The default value of  **Top** is -19, which means the active window is 19 points above the window display area. If you set the value to less than -19, part of the active window is hidden below the ribbon.

For the distance of the main window from the top of the screen, see the  **[Top](application-top-property-project.md)** property of the **Application** object.


