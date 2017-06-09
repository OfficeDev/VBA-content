---
title: GlowFormat.Creator Property (Word)
keywords: vbawd10.chm255853545
f1_keywords:
- vbawd10.chm255853545
ms.prod: word
api_name:
- Word.GlowFormat.Creator
ms.assetid: 37204a1d-2ac6-75fe-d843-1e91826e7ac1
ms.date: 06/08/2017
---


# GlowFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **GlowFormat** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[GlowFormat Object](glowformat-object-word.md)

