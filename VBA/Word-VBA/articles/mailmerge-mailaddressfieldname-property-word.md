---
title: MailMerge.MailAddressFieldName Property (Word)
keywords: vbawd10.chm153092105
f1_keywords:
- vbawd10.chm153092105
ms.prod: word
api_name:
- Word.MailMerge.MailAddressFieldName
ms.assetid: 729e6afa-26a6-75dd-78f8-9677aedfb2fa
ms.date: 06/08/2017
---


# MailMerge.MailAddressFieldName Property (Word)

Returns or sets the name of the field that contains e-mail addresses that are used when the mail merge destination is electronic mail. Read/write  **String** .


## Syntax

 _expression_ . **MailAddressFieldName**

 _expression_ An expression that returns a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example merges the document named "FormLetter.doc" with its attached data document and sends the results to the e-mail addresses stored in the Email merge field.


```vb
With Documents("FormLetter.doc").MailMerge 
 .MailAddressFieldName = "Email" 
 .MailSubject = "Amazing offer" 
 .Destination = wdSendToEmail 
 .Execute 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

