---
title: Application.MailLogoff Method (Excel)
keywords: vbaxl10.chm133157
f1_keywords:
- vbaxl10.chm133157
ms.prod: excel
api_name:
- Excel.Application.MailLogoff
ms.assetid: 5265e9c1-6c04-3591-7133-5274e5b56347
ms.date: 06/08/2017
---


# Application.MailLogoff Method (Excel)

Closes a MAPI mail session established by Microsoft Excel.


## Syntax

 _expression_ . **MailLogoff**

 _expression_ A variable that represents an **Application** object.


## Remarks

You cannot use this method to close or log off Microsoft Mail.


## Example

This example closes the established mail session, if there is one.


```vb
If Not IsNull(Application.MailSession) Then Application.MailLogoff
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

