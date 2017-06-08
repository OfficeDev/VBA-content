---
title: ChartArea.Creator Property (Word)
keywords: vbawd10.chm160039061
f1_keywords:
- vbawd10.chm160039061
ms.prod: word
api_name:
- Word.ChartArea.Creator
ms.assetid: dd2c0557-336b-ec22-fed4-dfa263d9fb52
ms.date: 06/08/2017
---


# ChartArea.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ChartArea](chartarea-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[ChartArea Object](chartarea-object-word.md)

