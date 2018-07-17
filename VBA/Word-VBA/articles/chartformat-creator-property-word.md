---
title: ChartFormat.Creator Property (Word)
keywords: vbawd10.chm203030677
f1_keywords:
- vbawd10.chm203030677
ms.prod: word
api_name:
- Word.ChartFormat.Creator
ms.assetid: a17057d9-8539-96e5-4f4a-222ccf13eaae
ms.date: 06/08/2017
---


# ChartFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ChartFormat](chartformat-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[ChartFormat Object](chartformat-object-word.md)

