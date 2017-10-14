---
title: ProofreadingErrors.Creator Property (Word)
keywords: vbawd10.chm163251177
f1_keywords:
- vbawd10.chm163251177
ms.prod: word
api_name:
- Word.ProofreadingErrors.Creator
ms.assetid: 230d4232-3171-0029-52ed-c8724e666fa8
ms.date: 06/08/2017
---


# ProofreadingErrors.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[ProofreadingErrors](proofreadingerrors-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[ProofreadingErrors Collection Object](proofreadingerrors-object-word.md)

