---
title: Axes.Creator Property (Word)
keywords: vbawd10.chm93388949
f1_keywords:
- vbawd10.chm93388949
ms.prod: word
api_name:
- Word.Axes.Creator
ms.assetid: 09557e5f-fd81-c7f4-a2a3-b842ef05a5d7
ms.date: 06/08/2017
---


# Axes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[Axes](axes-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[Axes Object](axes-object-word.md)

