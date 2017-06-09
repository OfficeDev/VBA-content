---
title: TaskRequestUpdateItem.BeforeAttachmentPreview Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.BeforeAttachmentPreview
ms.assetid: 3f071f28-40ba-53af-82de-23fff1b2a521
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.BeforeAttachmentPreview Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is previewed.


## Syntax

 _expression_ . **BeforeAttachmentPreview**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be previewed.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be previewed.|

## Remarks

This event occurs before an attachment is previewed, either from the attachment strip in the Reading Pane of the active explorer or from the active inspector.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

