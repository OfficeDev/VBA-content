---
title: RemoteItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.BeforeAttachmentWriteToTempFile
ms.assetid: fb309e7f-b8a6-b73c-de7a-77a15a70249d
ms.date: 06/08/2017
---


# RemoteItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_ . **BeforeAttachmentWriteToTempFile**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

