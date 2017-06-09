---
title: TaskItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.BeforeAttachmentWriteToTempFile
ms.assetid: 6f6acd79-afc2-7b40-60c9-770b8561b1a9
ms.date: 06/08/2017
---


# TaskItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_ . **BeforeAttachmentWriteToTempFile**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

