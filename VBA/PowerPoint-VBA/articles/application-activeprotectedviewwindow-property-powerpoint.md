---
title: Application.ActiveProtectedViewWindow Property (PowerPoint)
keywords: vbapp10.chm503014
f1_keywords:
- vbapp10.chm503014
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActiveProtectedViewWindow
ms.assetid: c0a7e748-d7fc-4a63-62b8-0eed5cf1c5b5
ms.date: 06/08/2017
---


# Application.ActiveProtectedViewWindow Property (PowerPoint)

Returns a  **[ProtectedViewWindow](protectedviewwindow-object-powerpoint.md)** object that represents the active **Protected View** window (the window on top). Read-only.


## Syntax

 _expression_. **ActiveProtectedViewWindow**

 _expression_ A variable that represents an **Application** object.


## Remarks

 **Nothing** if there are no **Protected View** windows open.


## Example

The following example displays the name ( **Caption** property) of the active **Protected View** window.


```vb
MsgBox "The name of the active Protected View window is " &; ActiveProtectedWindow.Caption
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

