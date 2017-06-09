---
title: Frames.Creator Property (Word)
keywords: vbawd10.chm153813993
f1_keywords:
- vbawd10.chm153813993
ms.prod: word
api_name:
- Word.Frames.Creator
ms.assetid: 403101cd-91c0-bb04-0b7e-a71117097391
ms.date: 06/08/2017
---


# Frames.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **Frames** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Frames Object](frames-object-word.md)

