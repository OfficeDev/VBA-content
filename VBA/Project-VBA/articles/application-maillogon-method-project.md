---
title: Application.MailLogon Method (Project)
ms.prod: project-server
api_name:
- Project.Application.MailLogon
ms.assetid: 0047a6ea-ea36-498c-e744-c4c88a08baae
ms.date: 06/08/2017
---


# Application.MailLogon Method (Project)

Logs on to a MAPI mail system and establishes a mail session. A mail session must be established before mail or document routing methods can be used.


## Syntax

 _expression_. **MailLogon**( ** _Name_**, ** _Password_**, ** _DownloadNewMail_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The mail account name.|
| _Password_|Optional|**String**|The mail account password.|
| _DownloadNewMail_|Optional|**Boolean**|**True** if new mail is downloaded immediately.|

## Remarks

Previously established mail sessions are logged off before an attempt is made to establish the new session. Omit both  _Name_ and _Password_ to use the default mail session for the system.


## Example

The following example logs on to the mail system and downloads any new mail.


```vb
Sub SessionLogon() 
 
 If IsNull(MailSession) Then 
 Application.MailLogon "oscarx", "mypassword", True 
 End If 
 
End Sub
```


