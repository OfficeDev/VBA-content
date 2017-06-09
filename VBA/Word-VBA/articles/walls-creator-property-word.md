---
title: Walls.Creator Property (Word)
keywords: vbawd10.chm25165973
f1_keywords:
- vbawd10.chm25165973
ms.prod: word
api_name:
- Word.Walls.Creator
ms.assetid: 1d47046e-3552-43d9-79f0-2317f8df380e
ms.date: 06/08/2017
---


# Walls.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[Walls](walls-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[Walls Object](walls-object-word.md)

