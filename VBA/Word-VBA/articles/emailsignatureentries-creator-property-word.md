---
title: EmailSignatureEntries.Creator Property (Word)
keywords: vbawd10.chm166003689
f1_keywords:
- vbawd10.chm166003689
ms.prod: word
api_name:
- Word.EmailSignatureEntries.Creator
ms.assetid: cca4c109-d018-50bf-72d0-0d84b1436aad
ms.date: 06/08/2017
---


# EmailSignatureEntries.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[EmailSignatureEntries](emailsignatureentries-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[EmailSignatureEntries Collection](emailsignatureentries-object-word.md)

