---
title: DataTable.Creator Property (Word)
keywords: vbawd10.chm46399637
f1_keywords:
- vbawd10.chm46399637
ms.prod: word
api_name:
- Word.DataTable.Creator
ms.assetid: 15f5bb8c-ac5a-5e51-45bb-1b163245e283
ms.date: 06/08/2017
---


# DataTable.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[DataTable](datatable-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[DataTable Object](datatable-object-word.md)

