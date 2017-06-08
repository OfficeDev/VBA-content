---
title: LegendKey.Creator Property (Word)
keywords: vbawd10.chm266207381
f1_keywords:
- vbawd10.chm266207381
ms.prod: word
api_name:
- Word.LegendKey.Creator
ms.assetid: ec1942b0-1ba3-cb55-1e0f-1bb8258f4810
ms.date: 06/08/2017
---


# LegendKey.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[LegendKey](legendkey-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[LegendKey Object](legendkey-object-word.md)

