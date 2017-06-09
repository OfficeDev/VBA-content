---
title: Points.Creator Property (Word)
keywords: vbawd10.chm10485909
f1_keywords:
- vbawd10.chm10485909
ms.prod: word
api_name:
- Word.Points.Creator
ms.assetid: b752ccbf-975a-bf99-4b24-52bb111ef251
ms.date: 06/08/2017
---


# Points.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[Points](points-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[Points Object](points-object-word.md)

