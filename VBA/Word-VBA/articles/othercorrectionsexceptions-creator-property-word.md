---
title: OtherCorrectionsExceptions.Creator Property (Word)
keywords: vbawd10.chm165610473
f1_keywords:
- vbawd10.chm165610473
ms.prod: word
api_name:
- Word.OtherCorrectionsExceptions.Creator
ms.assetid: b555cd72-95a8-edd9-a335-5885b85ef517
ms.date: 06/08/2017
---


# OtherCorrectionsExceptions.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[OtherCorrectionsExceptions](othercorrectionsexceptions-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[OtherCorrectionsExceptions Collection Object](othercorrectionsexceptions-object-word.md)

