---
title: Attachment.Delete Method (Publisher)
keywords: vbapb10.chm573441
f1_keywords:
- vbapb10.chm573441
ms.prod: publisher
api_name:
- Publisher.Attachment.Delete
ms.assetid: 935fa9e7-9d40-b820-e386-1a1960845da1
ms.date: 06/08/2017
---


# Attachment.Delete Method (Publisher)

Deletes an  **Attachment** object from the **Attachments** collection of an e-mail merge message.


## Syntax

 _expression_. **Delete**

 _expression_A variable that represents an  **Attachment** object.


## Remarks

The  **Delete** method performs an irreversible operation on the **Attachments** collection. It calls **IUnknown.Release** on the collection's reference to the **Attachment** object. If you have another reference to the attachment, you can still access its properties and methods, but you can never again associate it with any collection, because the **[Add](attachments-add-method-publisher.md)** method always creates a new object. Use the **Set** keyword to set your reference variable either to **Nothing** or to another attachment.

The final release of the  **Attachment** object takes place when you assign your reference variable to **Nothing**, or when you call  **Delete**, if you had no other reference. At this point the object is removed from memory. Attempting to gain access to a released object returns the Microsoft Collaboration Data Object error  **CdoE_INVALID_OBJECT**.

When you delete a member of a collection, the collection is immediately refreshed, meaning that its  **Count** property is reduced by one and its members are reindexed. To access the member that previously followed the deleted member in the collection, you must use its new index value.

To delete all attachments to the current e-mail merge message, use the  **[ClearAll](attachments-clearall-method-publisher.md)** method of the **Attachments** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to delete an attachment to the message in an e-mail merge. The code deletes the attachment at the first index position in the  **Attachments** collection and then prints the name of the deleted attachment and the number of current attachments to the message in the **Immediate** window.

Before running this code, ensure that there is at least one attachment to the current e-mail merge message.




```vb
Public Sub Delete_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Set pubAttachment = pubAttachments(1) 
 Debug.Print pubAttachments.Count 
 Debug.Print pubAttachment.Name 
 
 pubAttachment.Delete 
 
End Sub
```


## See also


#### Concepts


 [Attachment Object](attachment-object-publisher.md)

