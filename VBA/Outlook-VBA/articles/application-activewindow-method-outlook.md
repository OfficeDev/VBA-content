---
title: Application.ActiveWindow Method (Outlook)
keywords: vbaol11.chm726
f1_keywords:
- vbaol11.chm726
ms.prod: outlook
api_name:
- Outlook.Application.ActiveWindow
ms.assetid: 5f5b4e8b-61e4-417b-6b0c-14d1ccb41594
ms.date: 06/08/2017
---


# Application.ActiveWindow Method (Outlook)

Returns an object representing the current Microsoft Outlook window on the desktop, either an  **[Explorer](explorer-object-outlook.md)** or an **[Inspector](inspector-object-outlook.md)** object.


## Syntax

 _expression_ . **ActiveWindow**

 _expression_ A variable that represents an **Application** object.


### Return Value

An  **Object** that represents the current Outlook window on the desktop. Returns **Nothing** if no Outlook explorer or inspector is open.


## Example

This Microsoft Visual Basic for Applications (VBA) example minimizes the topmost Outlook window if it is an inspector window.


```vb
Sub MinimizeActiveWindow() 
 
 If TypeName(Application.ActiveWindow) = "Inspector" Then 
 
 Application.ActiveWindow.WindowState = olMinimized 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

