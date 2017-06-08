---
title: InlineShapes.Creator Property (Word)
keywords: vbawd10.chm162071529
f1_keywords:
- vbawd10.chm162071529
ms.prod: word
api_name:
- Word.InlineShapes.Creator
ms.assetid: 6f60b57b-12a8-997d-8043-ac33ab1e6840
ms.date: 06/08/2017
---


# InlineShapes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[InlineShapes](inlineshapes-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[InlineShapes Collection Object](inlineshapes-object-word.md)

