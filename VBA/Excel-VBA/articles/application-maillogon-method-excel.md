---
title: Application.MailLogon Method (Excel)
keywords: vbaxl10.chm133158
f1_keywords:
- vbaxl10.chm133158
ms.prod: excel
api_name:
- Excel.Application.MailLogon
ms.assetid: 0a6c8752-739d-b996-1426-4d3021ea5323
ms.date: 06/08/2017
---


# Application.MailLogon Method (Excel)

Logs in to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail isn't already running, you must use this method to establish a mail session before mail or document routing functions can be used.


## Syntax

 _expression_ . **MailLogon**( **_Name_** , **_Password_** , **_DownloadNewMail_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|The mail account name or Microsoft Exchange profile name. If this argument is omitted, the default mail account name is used.|
| _Password_|Optional| **Variant**|The mail account password. This argument is ignored in Microsoft Exchange.|
| _DownloadNewMail_|Optional| **Variant**| **True** to download new mail immediately.|

## Remarks

Microsoft Excel logs off any mail sessions it previously established before attempting to establish the new session.

To piggyback on the system default mail session, omit both the name and password parameters.


## Example

This example logs in to the default mail account.


```vb
If IsNull(Application.MailSession) Then 
 Application.MailLogon 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

