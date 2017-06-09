---
title: CustomProperties.Creator Property (Word)
keywords: vbawd10.chm232850409
f1_keywords:
- vbawd10.chm232850409
ms.prod: word
api_name:
- Word.CustomProperties.Creator
ms.assetid: 52c96a6e-b649-e861-fd69-055f526b6f7a
ms.date: 06/08/2017
---


# CustomProperties.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[CustomProperties](customproperties-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[CustomProperties Collection](customproperties-object-word.md)

