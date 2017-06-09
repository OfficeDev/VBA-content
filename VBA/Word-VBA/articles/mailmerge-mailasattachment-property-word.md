---
title: MailMerge.MailAsAttachment Property (Word)
keywords: vbawd10.chm153092104
f1_keywords:
- vbawd10.chm153092104
ms.prod: word
api_name:
- Word.MailMerge.MailAsAttachment
ms.assetid: ffa6505c-e14f-9315-0bc6-ff84ffb39931
ms.date: 06/08/2017
---


# MailMerge.MailAsAttachment Property (Word)

 **True** if the merge documents are sent as attachments when the mail merge destination is an e-mail message or a fax. Read/write **Boolean** .


## Syntax

 _expression_ . **MailAsAttachment**

 _expression_ An expression that returns a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example performs a mail merge operation and sends the merge results as attachments to e-mail messages. The e-mail addresses are stored in the MailAddress merge field.


```vb
With Documents("Main.doc").MailMerge 
 .MailAsAttachment = True 
 .Destination = wdSendToEmail 
 .MailSubject = "Special offer" 
 .MailAddressFieldName = "MailAddress" 
 .Execute 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

