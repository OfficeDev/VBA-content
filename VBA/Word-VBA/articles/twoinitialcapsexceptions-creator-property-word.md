---
title: TwoInitialCapsExceptions.Creator Property (Word)
keywords: vbawd10.chm155452393
f1_keywords:
- vbawd10.chm155452393
ms.prod: word
api_name:
- Word.TwoInitialCapsExceptions.Creator
ms.assetid: b3a43d7e-0deb-f15e-7faa-7de37c90d5b8
ms.date: 06/08/2017
---


# TwoInitialCapsExceptions.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TwoInitialCapsExceptions](twoinitialcapsexceptions-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TwoInitialCapsExceptions Collection Object](twoinitialcapsexceptions-object-word.md)

