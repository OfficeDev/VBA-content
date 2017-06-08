---
title: TextColumn.Creator Property (Word)
keywords: vbawd10.chm158598121
f1_keywords:
- vbawd10.chm158598121
ms.prod: word
api_name:
- Word.TextColumn.Creator
ms.assetid: 212586eb-93ac-e892-9890-310f18e2e6ca
ms.date: 06/08/2017
---


# TextColumn.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TextColumn](textcolumn-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TextColumn Object](textcolumn-object-word.md)

