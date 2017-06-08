---
title: ReflectionFormat.Creator Property (Word)
keywords: vbawd10.chm67044329
f1_keywords:
- vbawd10.chm67044329
ms.prod: word
api_name:
- Word.ReflectionFormat.Creator
ms.assetid: 3b6282d3-6ae4-e436-5f0b-fb1dff16c41d
ms.date: 06/08/2017
---


# ReflectionFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **ReflectionFormat** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[ReflectionFormat Object](reflectionformat-object-word.md)

