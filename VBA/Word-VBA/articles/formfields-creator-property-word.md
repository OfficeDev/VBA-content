---
title: FormFields.Creator Property (Word)
keywords: vbawd10.chm153682921
f1_keywords:
- vbawd10.chm153682921
ms.prod: word
api_name:
- Word.FormFields.Creator
ms.assetid: 32fa2979-4542-a1eb-3753-c38c3edffc35
ms.date: 06/08/2017
---


# FormFields.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[FormFields](formfields-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[FormFields Collection Object](formfields-object-word.md)

