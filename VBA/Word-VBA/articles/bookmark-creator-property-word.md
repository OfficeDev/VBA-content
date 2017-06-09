---
title: Bookmark.Creator Property (Word)
keywords: vbawd10.chm157811689
f1_keywords:
- vbawd10.chm157811689
ms.prod: word
api_name:
- Word.Bookmark.Creator
ms.assetid: 3994cad5-c492-8043-4fdf-d1900fdf08c4
ms.date: 06/08/2017
---


# Bookmark.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[Bookmark](bookmark-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Bookmark Object](bookmark-object-word.md)

