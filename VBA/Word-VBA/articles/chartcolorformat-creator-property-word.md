---
title: ChartColorFormat.Creator Property (Word)
keywords: vbawd10.chm12058773
f1_keywords:
- vbawd10.chm12058773
ms.prod: word
api_name:
- Word.ChartColorFormat.Creator
ms.assetid: 56389a3f-8633-ed9f-dd08-c495bf48cf5c
ms.date: 06/08/2017
---


# ChartColorFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ChartColorFormat](chartcolorformat-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[ChartColorFormat Object](chartcolorformat-object-word.md)

