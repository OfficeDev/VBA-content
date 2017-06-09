---
title: Application.Startup Event (Outlook)
keywords: vbaol11.chm433
f1_keywords:
- vbaol11.chm433
ms.prod: outlook
api_name:
- Outlook.Application.Startup
ms.assetid: d4724d96-2572-b1e3-e202-0bfffb5cf7d5
ms.date: 06/08/2017
---


# Application.Startup Event (Outlook)

Occurs when Microsoft Outlook is starting, but after all add-in programs have been loaded. 


## Syntax

 _expression_ . **Startup**

 _expression_ A variable that represents an **Application** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

An Outlook macro can use this event procedure to initialize itself when Outlook starts.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays a welcome message to the user and maximizes the Outlook explorer window when Outlook starts.


```vb
Private Sub Application_Startup() 
 
 MsgBox "Welcome, " &; Application.GetNamespace("MAPI").CurrentUser 
 
 Application.ActiveExplorer.WindowState = olMaximized 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

