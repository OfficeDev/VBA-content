---
title: SynonymInfo.Creator Property (Word)
keywords: vbawd10.chm161154025
f1_keywords:
- vbawd10.chm161154025
ms.prod: word
api_name:
- Word.SynonymInfo.Creator
ms.assetid: 04eb1a39-a345-9118-ddd5-5db6f062acf8
ms.date: 06/08/2017
---


# SynonymInfo.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)

