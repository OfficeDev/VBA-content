---
title: FreeformBuilder.Creator Property (Word)
keywords: vbawd10.chm164168681
f1_keywords:
- vbawd10.chm164168681
ms.prod: word
api_name:
- Word.FreeformBuilder.Creator
ms.assetid: f0e2b402-4de4-b864-ea35-8fc3c6e97a1e
ms.date: 06/08/2017
---


# FreeformBuilder.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[FreeformBuilder](freeformbuilder-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[FreeformBuilder Object](freeformbuilder-object-word.md)

