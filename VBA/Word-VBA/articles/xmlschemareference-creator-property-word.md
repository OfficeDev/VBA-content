---
title: XMLSchemaReference.Creator Property (Word)
keywords: vbawd10.chm32506857
f1_keywords:
- vbawd10.chm32506857
ms.prod: word
api_name:
- Word.XMLSchemaReference.Creator
ms.assetid: f2153a6e-0be9-2bf3-f2ba-3c21f99a7021
ms.date: 06/08/2017
---


# XMLSchemaReference.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[XMLSchemaReference](xmlschemareference-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[XMLSchemaReference Object](xmlschemareference-object-word.md)

