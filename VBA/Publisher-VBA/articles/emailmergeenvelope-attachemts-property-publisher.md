---
title: EmailMergeEnvelope.Attachemts Property (Publisher)
keywords: vbapb10.chm9043975
f1_keywords:
- vbapb10.chm9043975
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Attachemts
ms.assetid: 53948bf7-2727-7b9c-a645-c9b954d5e023
ms.date: 06/08/2017
---


# EmailMergeEnvelope.Attachemts Property (Publisher)

Gets the list of a merged e-mail message's attachments. Read-only.


## Syntax

 _expression_. **Attachemts**

 _expression_A variable that represents an  **EmailMergeEnvelope** object.


### Return Value

Attachments


## Remarks

To add attachments to a merged e-mail message, use the  **[Add](attachments-add-method-publisher.md)** method of the **[Attachment](attachment-object-publisher.md)** object. To remove an attachment, use the ** [Attachment.Delete](attachment-delete-method-publisher.md)** method; to remove all attachments, use the **[ClearAll](attachments-clearall-method-publisher.md)** method of the **[Attachments](attachments-object-publisher.md)** collection.


