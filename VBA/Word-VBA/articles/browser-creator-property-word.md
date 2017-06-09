---
title: Browser.Creator Property (Word)
keywords: vbawd10.chm154010601
f1_keywords:
- vbawd10.chm154010601
ms.prod: word
api_name:
- Word.Browser.Creator
ms.assetid: dd12021b-a90c-d24f-6556-01d3f5ebd582
ms.date: 06/08/2017
---


# Browser.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[Browser](browser-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Browser Object](browser-object-word.md)

