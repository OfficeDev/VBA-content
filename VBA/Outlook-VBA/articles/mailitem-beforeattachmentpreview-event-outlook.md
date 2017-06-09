---
title: MailItem.BeforeAttachmentPreview Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAttachmentPreview
ms.assetid: 279e1af4-38e1-d6b5-50a5-9ebd517826ae
ms.date: 06/08/2017
---


# MailItem.BeforeAttachmentPreview Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is previewed.


## Syntax

 _expression_ . **BeforeAttachmentPreview**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be previewed.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be previewed.|

## Remarks

This event occurs before an attachment is previewed, either from the attachment strip in the Reading Pane of the active explorer or from the active inspector.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

