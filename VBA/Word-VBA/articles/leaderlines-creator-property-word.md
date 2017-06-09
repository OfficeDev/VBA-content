---
title: LeaderLines.Creator Property (Word)
keywords: vbawd10.chm207749269
f1_keywords:
- vbawd10.chm207749269
ms.prod: word
api_name:
- Word.LeaderLines.Creator
ms.assetid: 2e23e29b-6008-d534-9160-10ec27c21b98
ms.date: 06/08/2017
---


# LeaderLines.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[LeaderLines](leaderlines-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[LeaderLines Object](leaderlines-object-word.md)

