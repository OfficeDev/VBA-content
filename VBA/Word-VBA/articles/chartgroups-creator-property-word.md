---
title: ChartGroups.Creator Property (Word)
keywords: vbawd10.chm77004949
f1_keywords:
- vbawd10.chm77004949
ms.prod: word
api_name:
- Word.ChartGroups.Creator
ms.assetid: 580937b3-8066-7208-ff98-f023dd30b713
ms.date: 06/08/2017
---


# ChartGroups.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ChartGroups](chartgroups-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[ChartGroups Object](chartgroups-object-word.md)

