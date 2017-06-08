---
title: SharingItem.BeforeAttachmentAdd Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeAttachmentAdd
ms.assetid: 84c14b49-a410-5e85-159d-b3f24a1dcad9
ms.date: 06/08/2017
---


# SharingItem.BeforeAttachmentAdd Event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

 _expression_ . **BeforeAttachmentAdd**( **_Attachment_** , **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

