---
title: DocumentItem.BeforeAttachmentAdd Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.BeforeAttachmentAdd
ms.assetid: cd440e8a-c79a-d1b4-9d03-940b2f3fa50b
ms.date: 06/08/2017
---


# DocumentItem.BeforeAttachmentAdd Event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

 _expression_ . **BeforeAttachmentAdd**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **DocumentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

