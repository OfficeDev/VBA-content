---
title: TableOfFigures.Creator Property (Word)
keywords: vbawd10.chm153158633
f1_keywords:
- vbawd10.chm153158633
ms.prod: word
api_name:
- Word.TableOfFigures.Creator
ms.assetid: c458a715-6d70-97f4-7a55-15b2735af5f9
ms.date: 06/08/2017
---


# TableOfFigures.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

