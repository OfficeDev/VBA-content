---
title: Lists.Creator Property (Word)
keywords: vbawd10.chm160629737
f1_keywords:
- vbawd10.chm160629737
ms.prod: word
api_name:
- Word.Lists.Creator
ms.assetid: 9c7a1ba2-479b-30af-1a9c-4f0f7fb73bf7
ms.date: 06/08/2017
---


# Lists.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Lists](lists-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Lists Collection Object](lists-object-word.md)

