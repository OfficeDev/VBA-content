---
title: TaskRequestDeclineItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.BeforeAttachmentWriteToTempFile
ms.assetid: c9564849-ecb2-a5a2-1c7e-f7cfea5ce34d
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_ . **BeforeAttachmentWriteToTempFile**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

