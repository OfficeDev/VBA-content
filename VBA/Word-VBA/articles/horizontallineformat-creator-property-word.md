---
title: HorizontalLineFormat.Creator Property (Word)
keywords: vbawd10.chm165544937
f1_keywords:
- vbawd10.chm165544937
ms.prod: word
api_name:
- Word.HorizontalLineFormat.Creator
ms.assetid: ecbfac32-9cf2-286d-4c7a-ca74056b8d40
ms.date: 06/08/2017
---


# HorizontalLineFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[HorizontalLineFormat](horizontallineformat-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[HorizontalLineFormat Object](horizontallineformat-object-word.md)

