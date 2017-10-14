---
title: Application.Quit Event (Outlook)
keywords: vbaol11.chm434
f1_keywords:
- vbaol11.chm434
ms.prod: outlook
api_name:
- Outlook.Application.Quit
ms.assetid: ecf0b50b-db6f-7eaf-90bd-bae942bf9287
ms.date: 06/08/2017
---


# Application.Quit Event (Outlook)

Occurs when Microsoft Outlook begins to close. 


## Syntax

 _expression_ . **Quit**

 _expression_ An expression that returns an **Application** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example displays a farewell message when Outlook exits. The sample code must be placed in a class module.


```vb
Private Sub Application_Quit() 
 
 MsgBox "Goodbye, " &; Application.GetNamespace("MAPI").CurrentUser 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

