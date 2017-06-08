---
title: TaskRequestItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.BeforeAttachmentRead
ms.assetid: 8d512d24-14e8-2c60-d70a-0f29ea24b618
ms.date: 06/08/2017
---


# TaskRequestItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

