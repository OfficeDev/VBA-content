---
title: MappedDataFields.Creator Property (Word)
keywords: vbawd10.chm135529449
f1_keywords:
- vbawd10.chm135529449
ms.prod: word
api_name:
- Word.MappedDataFields.Creator
ms.assetid: 1948ecf8-a42d-7a1b-16d2-808caa53dd9a
ms.date: 06/08/2017
---


# MappedDataFields.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[MappedDataFields](mappeddatafields-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[MappedDataFields Collection](mappeddatafields-object-word.md)

