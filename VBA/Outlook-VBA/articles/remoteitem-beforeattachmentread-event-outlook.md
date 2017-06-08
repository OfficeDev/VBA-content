---
title: RemoteItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.BeforeAttachmentRead
ms.assetid: 739b8606-3e3a-1445-6355-896a6e897a6f
ms.date: 06/08/2017
---


# RemoteItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

