---
title: AutoCorrectEntries.Creator Property (Word)
keywords: vbawd10.chm155714537
f1_keywords:
- vbawd10.chm155714537
ms.prod: word
api_name:
- Word.AutoCorrectEntries.Creator
ms.assetid: 65cff427-3520-863d-ddbd-5c1e83a9fe43
ms.date: 06/08/2017
---


# AutoCorrectEntries.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[AutoCorrectEntries](autocorrectentries-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[AutoCorrectEntries Collection Object](autocorrectentries-object-word.md)

