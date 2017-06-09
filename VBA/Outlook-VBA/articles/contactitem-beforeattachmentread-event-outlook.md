---
title: ContactItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.BeforeAttachmentRead
ms.assetid: ba862dea-f2e1-a864-f6c3-a8987c28bfcf
ms.date: 06/08/2017
---


# ContactItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

