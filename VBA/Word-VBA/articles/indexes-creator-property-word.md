---
title: Indexes.Creator Property (Word)
keywords: vbawd10.chm159122409
f1_keywords:
- vbawd10.chm159122409
ms.prod: word
api_name:
- Word.Indexes.Creator
ms.assetid: 88fed4ac-033b-a33f-0355-c750fcea0783
ms.date: 06/08/2017
---


# Indexes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[Indexes](indexes-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Indexes Collection Object](indexes-object-word.md)

