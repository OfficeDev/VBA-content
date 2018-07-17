---
title: Column.Creator Property (Word)
keywords: vbawd10.chm156173289
f1_keywords:
- vbawd10.chm156173289
ms.prod: word
api_name:
- Word.Column.Creator
ms.assetid: cadc230a-006c-6b2c-4ae0-115652f5280a
ms.date: 06/08/2017
---


# Column.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Column](column-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Column Object](column-object-word.md)

