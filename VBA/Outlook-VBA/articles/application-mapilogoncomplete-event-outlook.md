---
title: Application.MAPILogonComplete Event (Outlook)
keywords: vbaol11.chm437
f1_keywords:
- vbaol11.chm437
ms.prod: outlook
api_name:
- Outlook.Application.MAPILogonComplete
ms.assetid: db6f7cf8-2a45-560f-f592-613de86e08e2
ms.date: 06/08/2017
---


# Application.MAPILogonComplete Event (Outlook)

Occurs after the user has logged onto the system.


## Syntax

 _expression_ . **MAPILogonComplete**

 _expression_ A variable that represents an **Application** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays a message after the user has logged onto the system.


```vb
Private Sub Application_MAPILogonComplete() 
 
'Occurs when a user has logged on 
 
 
 
 MsgBox "Logon complete." 
 
 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

