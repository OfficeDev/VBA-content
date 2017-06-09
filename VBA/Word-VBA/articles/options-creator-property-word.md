---
title: Options.Creator Property (Word)
keywords: vbawd10.chm162989033
f1_keywords:
- vbawd10.chm162989033
ms.prod: word
api_name:
- Word.Options.Creator
ms.assetid: db5a4d71-9791-9dbb-7411-99bc9b1a5480
ms.date: 06/08/2017
---


# Options.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Options Object](options-object-word.md)

