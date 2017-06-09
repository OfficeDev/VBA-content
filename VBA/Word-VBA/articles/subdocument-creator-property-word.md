---
title: Subdocument.Creator Property (Word)
keywords: vbawd10.chm159974377
f1_keywords:
- vbawd10.chm159974377
ms.prod: word
api_name:
- Word.Subdocument.Creator
ms.assetid: 9b602f8e-433c-4679-cea5-37f6eea5f62d
ms.date: 06/08/2017
---


# Subdocument.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Subdocument](subdocument-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Subdocument Object](subdocument-object-word.md)

