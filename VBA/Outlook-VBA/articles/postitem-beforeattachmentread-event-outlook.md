---
title: PostItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.BeforeAttachmentRead
ms.assetid: c4e83a89-5ae9-ece3-b884-8f19adbdcc40
ms.date: 06/08/2017
---


# PostItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an **[Attachment](attachment-object-outlook.md)** object.


## Syntax

_expression_. **BeforeAttachmentRead** (**_Attachment_**, **_Cancel_**)

_expression_ A variable that represents a **PostItem** object.


### Parameters

|Name|Required/Optional|Data Type|Description|
|:-----|:-----|:-----|:-----|
|_Attachment_|Required|**Attachment**|The **Attachment** to be read.|
|_Cancel_|Required|**Boolean**|Set to **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

<br/>

## See also

#### Concepts

- [PostItem Object](postitem-object-outlook.md)

