---
title: ConditionalStyle.Creator Property (Word)
keywords: vbawd10.chm91030505
f1_keywords:
- vbawd10.chm91030505
ms.prod: word
api_name:
- Word.ConditionalStyle.Creator
ms.assetid: f0311e68-ce40-86fd-1009-adf361a072e0
ms.date: 06/08/2017
---


# ConditionalStyle.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[ConditionalStyle](conditionalstyle-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[ConditionalStyle Object](conditionalstyle-object-word.md)

