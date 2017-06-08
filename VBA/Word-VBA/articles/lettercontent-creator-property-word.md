---
title: LetterContent.Creator Property (Word)
keywords: vbawd10.chm161547241
f1_keywords:
- vbawd10.chm161547241
ms.prod: word
api_name:
- Word.LetterContent.Creator
ms.assetid: b2bee17a-490e-ebd5-5e3b-62e154d30a31
ms.date: 06/08/2017
---


# LetterContent.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

