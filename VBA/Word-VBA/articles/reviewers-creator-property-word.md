---
title: Reviewers.Creator Property (Word)
keywords: vbawd10.chm211420137
f1_keywords:
- vbawd10.chm211420137
ms.prod: word
api_name:
- Word.Reviewers.Creator
ms.assetid: 4a77f3a3-18ab-1d7a-ba8d-b773c1e6bc91
ms.date: 06/08/2017
---


# Reviewers.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[Reviewers](reviewers-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Reviewers Collection](reviewers-object-word.md)

