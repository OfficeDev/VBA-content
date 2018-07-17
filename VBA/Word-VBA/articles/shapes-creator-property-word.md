---
title: Shapes.Creator Property (Word)
keywords: vbawd10.chm161423169
f1_keywords:
- vbawd10.chm161423169
ms.prod: word
api_name:
- Word.Shapes.Creator
ms.assetid: 6eb44108-1323-ee89-6ffe-27aa0c6c7213
ms.date: 06/08/2017
---


# Shapes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

