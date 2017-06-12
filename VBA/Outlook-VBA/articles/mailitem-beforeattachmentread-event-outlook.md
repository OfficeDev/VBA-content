---
title: MailItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAttachmentRead
ms.assetid: 00d35fff-b1d2-0da2-7315-a9fce2f28e80
ms.date: 06/08/2017
---


# MailItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

