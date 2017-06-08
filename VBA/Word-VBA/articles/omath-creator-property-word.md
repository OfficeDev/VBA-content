---
title: OMath.Creator Property (Word)
keywords: vbawd10.chm176357477
f1_keywords:
- vbawd10.chm176357477
ms.prod: word
api_name:
- Word.OMath.Creator
ms.assetid: 7c009fb5-8b5d-e745-de33-7b6f6387765d
ms.date: 06/08/2017
---


# OMath.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[OMath](omath-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[OMath Object](omath-object-word.md)

