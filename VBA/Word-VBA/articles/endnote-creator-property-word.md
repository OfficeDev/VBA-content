---
title: Endnote.Creator Property (Word)
keywords: vbawd10.chm155059177
f1_keywords:
- vbawd10.chm155059177
ms.prod: word
api_name:
- Word.Endnote.Creator
ms.assetid: 673d007e-fe72-cc7d-e0eb-25e533b43f98
ms.date: 06/08/2017
---


# Endnote.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[Endnote](endnote-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Endnote Object](endnote-object-word.md)

