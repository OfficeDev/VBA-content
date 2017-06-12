---
title: Series.Creator Property (Word)
keywords: vbawd10.chm123732117
f1_keywords:
- vbawd10.chm123732117
ms.prod: word
api_name:
- Word.Series.Creator
ms.assetid: 640e4150-6aa8-1001-de42-c2fbe5f94460
ms.date: 06/08/2017
---


# Series.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[Series Object](series-object-word.md)

