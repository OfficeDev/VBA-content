---
title: ChartTitle.Creator Property (Word)
keywords: vbawd10.chm65274005
f1_keywords:
- vbawd10.chm65274005
ms.prod: word
api_name:
- Word.ChartTitle.Creator
ms.assetid: ff16dcb5-faac-edc7-21a2-631dd09cb12f
ms.date: 06/08/2017
---


# ChartTitle.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ChartTitle](charttitle-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[ChartTitle Object](charttitle-object-word.md)

