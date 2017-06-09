---
title: HangulAndAlphabetException.Creator Property (Word)
keywords: vbawd10.chm164758505
f1_keywords:
- vbawd10.chm164758505
ms.prod: word
api_name:
- Word.HangulAndAlphabetException.Creator
ms.assetid: 77b08c7a-411d-d7d7-a4fa-57c750d3b848
ms.date: 06/08/2017
---


# HangulAndAlphabetException.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[HangulAndAlphabetException](hangulandalphabetexception-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[HangulAndAlphabetException Object](hangulandalphabetexception-object-word.md)

