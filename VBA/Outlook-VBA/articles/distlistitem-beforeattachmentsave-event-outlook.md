---
title: DistListItem.BeforeAttachmentSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeAttachmentSave
ms.assetid: 5c96c5df-bb29-c2d5-1f80-5e4cd426f7b7
ms.date: 06/08/2017
---


# DistListItem.BeforeAttachmentSave Event (Outlook)

Occurs just before an attachment is saved.


## Syntax

 _expression_ . **BeforeAttachmentSave**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be saved.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed and the attachment is not changed.|

## Remarks

This event corresponds to when attachments are saved to the messaging store. The  **BeforeAttachmentSave** event occurs just before an attachment is saved when an item is saved. If a user edits an attachment and then saves those changes, the **BeforeAttachmentSave** event will not occur at that time; instead it will occur when the item itself is later saved. It also does not occur when the attachment is saved on the hard disk using the **SaveAsFile** method.

In VBScript, if you set the return value of this function to  **False** , the save operation is cancelled and the attachment is not changed.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

