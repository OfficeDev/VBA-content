---
title: ContactItem.BeforeAttachmentAdd Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.BeforeAttachmentAdd
ms.assetid: d0c0bfd1-5d18-759c-0131-c78e45982b18
ms.date: 06/08/2017
---


# ContactItem.BeforeAttachmentAdd Event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

 _expression_ . **BeforeAttachmentAdd**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

