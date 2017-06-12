---
title: Frame.Creator Property (Word)
keywords: vbawd10.chm153748457
f1_keywords:
- vbawd10.chm153748457
ms.prod: word
api_name:
- Word.Frame.Creator
ms.assetid: 0170c463-844d-46e0-ff6a-2db489545053
ms.date: 06/08/2017
---


# Frame.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Frame](frame-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Frame Object](frame-object-word.md)

