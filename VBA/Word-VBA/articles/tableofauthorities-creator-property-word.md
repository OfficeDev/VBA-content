---
title: TableOfAuthorities.Creator Property (Word)
keywords: vbawd10.chm152110057
f1_keywords:
- vbawd10.chm152110057
ms.prod: word
api_name:
- Word.TableOfAuthorities.Creator
ms.assetid: 22eda540-1af4-af1e-ef99-db471774f383
ms.date: 06/08/2017
---


# TableOfAuthorities.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TableOfAuthorities](tableofauthorities-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

