---
title: Characters.Creator Property (Word)
keywords: vbawd10.chm157090793
f1_keywords:
- vbawd10.chm157090793
ms.prod: word
api_name:
- Word.Characters.Creator
ms.assetid: d8bed9e7-237a-4049-79d1-1d68cc9ca0f1
ms.date: 06/08/2017
---


# Characters.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[Characters](characters-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Characters Collection Object](characters-object-word.md)

