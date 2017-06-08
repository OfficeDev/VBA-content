---
title: FileConverter.Creator Property (Word)
keywords: vbawd10.chm161022953
f1_keywords:
- vbawd10.chm161022953
ms.prod: word
api_name:
- Word.FileConverter.Creator
ms.assetid: c8015ff2-a16a-19c9-25b7-dd16fcf7220b
ms.date: 06/08/2017
---


# FileConverter.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

