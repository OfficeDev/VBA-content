---
title: CoAuthor.Creator Property (Word)
keywords: vbawd10.chm81069033
f1_keywords:
- vbawd10.chm81069033
ms.prod: word
api_name:
- Word.CoAuthor.Creator
ms.assetid: 30b59f4d-7d95-409d-463f-53c1492bf370
ms.date: 06/08/2017
---


# CoAuthor.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **CoAuthor** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD." This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[CoAuthor Object](coauthor-object-word.md)

