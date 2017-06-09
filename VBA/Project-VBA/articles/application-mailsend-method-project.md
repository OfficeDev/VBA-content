---
title: Application.MailSend Method (Project)
keywords: vbapj.chm120
f1_keywords:
- vbapj.chm120
ms.prod: project-server
api_name:
- Project.Application.MailSend
ms.assetid: 250c7eed-2bfa-f80f-13d1-c7ca8d6453d1
ms.date: 06/08/2017
---


# Application.MailSend Method (Project)

Sends a mail message.


## Syntax

 _expression_. **MailSend**( ** _To_**, ** _Cc_**, ** _Subject_**, ** _Body_**, ** _Enclosures_**, ** _IncludeDocument_**, ** _ReturnReceipt_**, ** _Bcc_**, ** _Urgent_**, ** _SaveCopy_**, ** _AddRecipient_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _To_|Optional|**String**|The user names of the primary recipients of the message, separated by commas.|
| _Cc_|Optional|**String**|The user names of the secondary recipients of the message, separated by commas.|
| _Subject_|Optional|**String**|The subject of the message.|
| _Body_|Optional|**String**|The main text of the message.|
| _Enclosures_|Optional|**String**|The file names of one or more files to include with the message. Use the list separator character to separate multiple file names. Do not add space between the list separator and the file name.|
| _IncludeDocument_|Optional|**Boolean**|**True** if the active project is included in the message. The default value is **True**.|
| _ReturnReceipt_|Optional|**Boolean**|**True** if a message is sent to the sender when the recipient opens the message. The default value is **False**.|
| _Bcc_|Optional|**String**|The user names of the message recipients which are not displayed, separated by semicolons. This argument is only supported in Microsoft Project for the Macintosh version 4.0|
| _Urgent_|Optional|**Boolean**|**True** if the message is given a high priority. This argument is only supported in Microsoft Project for the Macintosh version 4.0.|
| _SaveCopy_|Optional|**Boolean**|**True** if a copy of the message is saved in the SentItems folder. This argument is only supported in Microsoft Project for the Macintosh version 4.0.|
| _AddRecipient_|Optional|**Boolean**|**True** if recipients of the message are added to a personal address book. This argument is only supported in Microsoft Project for the Macintosh version 4.0.|

### Return Value

 **Boolean**


## Remarks

If the  **MailSend** method is used without specifying any arguments and there are no existing routing slips, a standard compose mail window appears with the active project as an embedded object. Otherwise, using the **MailSend** method without specifying any arguments prompts whether or not to use the routing slip.


