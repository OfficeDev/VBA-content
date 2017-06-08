---
title: Windows.Creator Property (Word)
keywords: vbawd10.chm157352937
f1_keywords:
- vbawd10.chm157352937
ms.prod: word
api_name:
- Word.Windows.Creator
ms.assetid: 6dfc07a8-e41a-de81-cfeb-6c0dff3d0a4b
ms.date: 06/08/2017
---


# Windows.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Windows](windows-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Windows Collection Object](windows-object-word.md)

