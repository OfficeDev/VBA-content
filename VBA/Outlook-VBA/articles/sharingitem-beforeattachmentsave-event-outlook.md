---
title: SharingItem.BeforeAttachmentSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeAttachmentSave
ms.assetid: ec6c8b9f-759b-df04-c3df-8e977df457a5
ms.date: 06/08/2017
---


# SharingItem.BeforeAttachmentSave Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read.


## Syntax

 _expression_ . **BeforeAttachmentSave**( **_Attachment_** , **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be saved.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be saved.|

## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

