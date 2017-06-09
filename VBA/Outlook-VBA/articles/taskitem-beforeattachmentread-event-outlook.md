---
title: TaskItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.BeforeAttachmentRead
ms.assetid: 298eaece-9633-637b-3055-572d77fa3811
ms.date: 06/08/2017
---


# TaskItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

