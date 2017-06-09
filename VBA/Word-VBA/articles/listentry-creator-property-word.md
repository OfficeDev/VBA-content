---
title: ListEntry.Creator Property (Word)
keywords: vbawd10.chm153289705
f1_keywords:
- vbawd10.chm153289705
ms.prod: word
api_name:
- Word.ListEntry.Creator
ms.assetid: c7120694-0e9b-6df8-7ce0-d7eb7fa6c456
ms.date: 06/08/2017
---


# ListEntry.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[ListEntry](listentry-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[ListEntry Object](listentry-object-word.md)

