---
title: TabStop.Creator Property (Word)
keywords: vbawd10.chm156500969
f1_keywords:
- vbawd10.chm156500969
ms.prod: word
api_name:
- Word.TabStop.Creator
ms.assetid: 5a8f0108-92d2-a6de-fb05-86da24bd157c
ms.date: 06/08/2017
---


# TabStop.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TabStop](tabstop-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TabStop Object](tabstop-object-word.md)

