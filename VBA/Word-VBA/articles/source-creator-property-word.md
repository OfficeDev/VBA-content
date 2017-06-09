---
title: Source.Creator Property (Word)
keywords: vbawd10.chm140836965
f1_keywords:
- vbawd10.chm140836965
ms.prod: word
api_name:
- Word.Source.Creator
ms.assetid: 19a8a10c-8b2f-3990-819d-8c06229e64a2
ms.date: 06/08/2017
---


# Source.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[Source](source-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Source Object](source-object-word.md)

