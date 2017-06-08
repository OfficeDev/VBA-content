---
title: Task.Creator Property (Word)
keywords: vbawd10.chm159515625
f1_keywords:
- vbawd10.chm159515625
ms.prod: word
api_name:
- Word.Task.Creator
ms.assetid: d0273cd1-fb98-dc8d-bf4a-bfcca4d053b1
ms.date: 06/08/2017
---


# Task.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Task](task-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Task Object](task-object-word.md)

