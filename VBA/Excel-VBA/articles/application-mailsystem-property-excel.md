---
title: Application.MailSystem Property (Excel)
keywords: vbaxl10.chm133160
f1_keywords:
- vbaxl10.chm133160
ms.prod: excel
api_name:
- Excel.Application.MailSystem
ms.assetid: df7b1238-bdf5-d9f8-9f50-585b489fd8a8
ms.date: 06/08/2017
---


# Application.MailSystem Property (Excel)

Returns the mail system that's installed on the host machine. Read-only  **[XlMailSystem](xlmailsystem-enumeration-excel.md)** .


## Syntax

 _expression_ . **MailSystem**

 _expression_ A variable that represents an **Application** object.


## Remarks





| **XlMailSystem** can be one of these **XlMailSystem** constants.|
| **xlMAPI**|
| **xlNoMailSystem**|
| **xlPowerTalk**|

## Example

This example displays the name of the mail system that's installed on the computer.


```vb
Select Case Application.MailSystem 
 Case xlMAPI 
 MsgBox "Mail system is Microsoft Mail" 
 Case xlPowerTalk 
 MsgBox "Mail system is PowerTalk" 
 Case xlNoMailSystem 
 MsgBox "No mail system installed" 
End Select
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

