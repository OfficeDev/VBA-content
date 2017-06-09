---
title: Cells.Creator Property (Word)
keywords: vbawd10.chm155845609
f1_keywords:
- vbawd10.chm155845609
ms.prod: word
api_name:
- Word.Cells.Creator
ms.assetid: 5113f3bd-2ac3-4ba3-5ab4-321ae6917eb2
ms.date: 06/08/2017
---


# Cells.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[Cells](cells-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

