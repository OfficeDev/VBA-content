---
title: HeadingStyle.Creator Property (Word)
keywords: vbawd10.chm160105449
f1_keywords:
- vbawd10.chm160105449
ms.prod: word
api_name:
- Word.HeadingStyle.Creator
ms.assetid: 2ae0b124-49ff-03b7-d11b-91b2115da057
ms.date: 06/08/2017
---


# HeadingStyle.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[HeadingStyle](headingstyle-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[HeadingStyle Object](headingstyle-object-word.md)

