---
title: TextColumns.Creator Property (Word)
keywords: vbawd10.chm158532585
f1_keywords:
- vbawd10.chm158532585
ms.prod: word
api_name:
- Word.TextColumns.Creator
ms.assetid: 312a79ae-6dd3-e616-d643-2f55202beefe
ms.date: 06/08/2017
---


# TextColumns.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TextColumns](textcolumns-objectword.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

