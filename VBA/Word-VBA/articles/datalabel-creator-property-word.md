---
title: DataLabel.Creator Property (Word)
keywords: vbawd10.chm233898133
f1_keywords:
- vbawd10.chm233898133
ms.prod: word
api_name:
- Word.DataLabel.Creator
ms.assetid: 3e261a34-9826-9c8e-5f5f-6fdd1101e9db
ms.date: 06/08/2017
---


# DataLabel.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[DataLabel Object](datalabel-object-word.md)

