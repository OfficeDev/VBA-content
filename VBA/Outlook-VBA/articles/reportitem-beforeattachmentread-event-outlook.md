---
title: ReportItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.BeforeAttachmentRead
ms.assetid: 65377c41-b51a-779c-9892-a61cc6e9b9da
ms.date: 06/08/2017
---


# ReportItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

