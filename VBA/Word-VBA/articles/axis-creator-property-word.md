---
title: Axis.Creator Property (Word)
keywords: vbawd10.chm113049749
f1_keywords:
- vbawd10.chm113049749
ms.prod: word
api_name:
- Word.Axis.Creator
ms.assetid: c7015ed2-d78f-4eb7-477c-11e896a7f37f
ms.date: 06/08/2017
---


# Axis.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[Axis Object](axis-object-word.md)

