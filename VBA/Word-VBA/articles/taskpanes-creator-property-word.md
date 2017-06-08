---
title: TaskPanes.Creator Property (Word)
ms.prod: word
api_name:
- Word.TaskPanes.Creator
ms.assetid: e94b0c6c-90a6-e221-2d56-966a197056bf
ms.date: 06/08/2017
---


# TaskPanes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TaskPanes](taskpanes-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TaskPanes Collection](taskpanes-object-word.md)

