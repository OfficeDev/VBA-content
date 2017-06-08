---
title: Application.Creator Property (Word)
keywords: vbawd10.chm158335977
f1_keywords:
- vbawd10.chm158335977
ms.prod: word
api_name:
- Word.Application.Creator
ms.assetid: 6afdfc30-5021-7b09-a148-48db16d5efbd
ms.date: 06/08/2017
---


# Application.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Application Object](application-object-word.md)

