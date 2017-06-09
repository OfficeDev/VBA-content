---
title: DistListItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeAttachmentWriteToTempFile
ms.assetid: 565c9c1d-cbab-1ac2-de79-3a15532c4f76
ms.date: 06/08/2017
---


# DistListItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_ . **BeforeAttachmentWriteToTempFile**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

