---
title: Zooms.Creator Property (Word)
keywords: vbawd10.chm161940457
f1_keywords:
- vbawd10.chm161940457
ms.prod: word
api_name:
- Word.Zooms.Creator
ms.assetid: 1434fc58-0bd2-c444-fb09-1cc430b4acf8
ms.date: 06/08/2017
---


# Zooms.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Zooms](zooms-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Zooms Collection Object](zooms-object-word.md)

