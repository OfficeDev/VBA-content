---
title: HeadingStyles.Creator Property (Word)
keywords: vbawd10.chm160039913
f1_keywords:
- vbawd10.chm160039913
ms.prod: word
api_name:
- Word.HeadingStyles.Creator
ms.assetid: 2faa2928-740a-2ecb-326b-241a82bb973e
ms.date: 06/08/2017
---


# HeadingStyles.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[HeadingStyles](headingstyles-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[HeadingStyles Collection Object](headingstyles-object-word.md)

