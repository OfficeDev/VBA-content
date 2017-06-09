---
title: SharingItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeAttachmentWriteToTempFile
ms.assetid: 85a7ac8e-94e2-1248-0d22-1ca8565c9530
ms.date: 06/08/2017
---


# SharingItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_ . **BeforeAttachmentWriteToTempFile**( **_Attachment_** , **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

