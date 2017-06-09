---
title: JournalItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.BeforeAttachmentRead
ms.assetid: a6200602-7939-9abb-d4f8-c7b1513325c8
ms.date: 06/08/2017
---


# JournalItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **JournalItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

