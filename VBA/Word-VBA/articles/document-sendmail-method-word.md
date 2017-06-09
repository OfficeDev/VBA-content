---
title: Document.SendMail Method (Word)
keywords: vbawd10.chm158007406
f1_keywords:
- vbawd10.chm158007406
ms.prod: word
api_name:
- Word.Document.SendMail
ms.assetid: 7e47982f-2c8f-e76b-d790-9c4e72d5110b
ms.date: 06/08/2017
---


# Document.SendMail Method (Word)

Opens a message window for sending the specified document through Microsoft Exchange.


## Syntax

 _expression_ . **SendMail**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use the  **SendMailAttach** property to control whether the document is sent as text in the message window or as an attachment.


## Example

This example sends the active document as an attachment to a mail message.


```vb
Options.SendMailAttach = True 
ActiveDocument.SendMail
```


## See also


#### Concepts


[Document Object](document-object-word.md)

