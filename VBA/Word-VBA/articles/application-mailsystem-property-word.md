---
title: Application.MailSystem Property (Word)
keywords: vbawd10.chm158335080
f1_keywords:
- vbawd10.chm158335080
ms.prod: word
api_name:
- Word.Application.MailSystem
ms.assetid: d8f97baa-2c50-c2af-0e50-f2de5d017b62
ms.date: 06/08/2017
---


# Application.MailSystem Property (Word)

Returns the mail system (or systems) installed on the host computer. Read-only  **WdMailSystem** .


## Syntax

 _expression_ . **MailSystem**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Some of the  **WdMailSystem** constants are available only in Microsoft Office Macintosh Edition. For additional information about these constants, consult the language reference Help included with Microsoft Office Macintosh Edition.


## Example

This example displays a message that indicates whether a mail system is installed on the computer.


```vb
ms = Application.MailSystem 
If ms <> wdNoMailSystem Then 
 MsgBox "This computer has a mail system installed." 
Else 
 MsgBox "This computer has no mail system installed." 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

