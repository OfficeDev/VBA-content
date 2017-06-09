---
title: TickLabels.Creator Property (Word)
keywords: vbawd10.chm167051413
f1_keywords:
- vbawd10.chm167051413
ms.prod: word
api_name:
- Word.TickLabels.Creator
ms.assetid: 854570ae-1e01-7b32-8c2d-8643c8912b82
ms.date: 06/08/2017
---


# TickLabels.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[TickLabels](ticklabels-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[TickLabels Object](ticklabels-object-word.md)

