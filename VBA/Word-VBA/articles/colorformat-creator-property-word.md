---
title: ColorFormat.Creator Property (Word)
keywords: vbawd10.chm163972073
f1_keywords:
- vbawd10.chm163972073
ms.prod: word
api_name:
- Word.ColorFormat.Creator
ms.assetid: 5a16e61d-2469-6e28-851a-f508ac0ce488
ms.date: 06/08/2017
---


# ColorFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ColorFormat](colorformat-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[ColorFormat Object](colorformat-object-word.md)

