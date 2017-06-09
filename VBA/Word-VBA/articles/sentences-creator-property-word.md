---
title: Sentences.Creator Property (Word)
keywords: vbawd10.chm156959721
f1_keywords:
- vbawd10.chm156959721
ms.prod: word
api_name:
- Word.Sentences.Creator
ms.assetid: 69465368-9258-cfc2-f469-69b27940e24e
ms.date: 06/08/2017
---


# Sentences.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Sentences](sentences-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Sentences Collection Object](sentences-object-word.md)

