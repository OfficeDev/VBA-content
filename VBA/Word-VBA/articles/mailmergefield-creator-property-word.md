---
title: MailMergeField.Creator Property (Word)
keywords: vbawd10.chm152962025
f1_keywords:
- vbawd10.chm152962025
ms.prod: word
api_name:
- Word.MailMergeField.Creator
ms.assetid: 7f7ac974-8233-b23d-72d8-b93d01660a8c
ms.date: 06/08/2017
---


# MailMergeField.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[MailMergeField](mailmergefield-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[MailMergeField Object](mailmergefield-object-word.md)

