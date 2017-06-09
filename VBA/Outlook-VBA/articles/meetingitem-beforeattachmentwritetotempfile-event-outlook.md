---
title: MeetingItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.BeforeAttachmentWriteToTempFile
ms.assetid: 26bbc5fc-4a65-101b-9693-f8d9ed9421c9
ms.date: 06/08/2017
---


# MeetingItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_ . **BeforeAttachmentWriteToTempFile**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **MeetingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

