---
title: ContactItem.BeforeAttachmentPreview Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.BeforeAttachmentPreview
ms.assetid: 7451778c-801a-15a9-203d-1a1c61ebc155
ms.date: 06/08/2017
---


# ContactItem.BeforeAttachmentPreview Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is previewed.


## Syntax

 _expression_ . **BeforeAttachmentPreview**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be previewed.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be previewed.|

## Remarks

This event occurs before an attachment is previewed, either from the attachment strip in the Reading Pane of the active explorer or from the active inspector.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

