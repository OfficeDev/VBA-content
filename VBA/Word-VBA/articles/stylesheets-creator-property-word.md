---
title: StyleSheets.Creator Property (Word)
keywords: vbawd10.chm209585129
f1_keywords:
- vbawd10.chm209585129
ms.prod: word
api_name:
- Word.StyleSheets.Creator
ms.assetid: ac2b45fd-3871-e134-4531-fb9c81fef4c4
ms.date: 06/08/2017
---


# StyleSheets.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[StyleSheets](stylesheets-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[StyleSheets Collection](stylesheets-object-word.md)

