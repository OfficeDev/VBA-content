---
title: Dialogs.Creator Property (Word)
keywords: vbawd10.chm152044521
f1_keywords:
- vbawd10.chm152044521
ms.prod: word
api_name:
- Word.Dialogs.Creator
ms.assetid: f5328ecf-d3de-85e4-cf00-b7fea51bf667
ms.date: 06/08/2017
---


# Dialogs.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Dialogs](dialogs-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Dialogs Collection Object](dialogs-object-word.md)

