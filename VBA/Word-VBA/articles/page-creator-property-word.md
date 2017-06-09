---
title: Page.Creator Property (Word)
keywords: vbawd10.chm11076585
f1_keywords:
- vbawd10.chm11076585
ms.prod: word
api_name:
- Word.Page.Creator
ms.assetid: 9f34c6ef-12d7-f494-095a-f9b59b696e98
ms.date: 06/08/2017
---


# Page.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Page](page-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Page Object](page-object-word.md)

