---
title: AutoCaption.Creator Property (Word)
keywords: vbawd10.chm159056873
f1_keywords:
- vbawd10.chm159056873
ms.prod: word
api_name:
- Word.AutoCaption.Creator
ms.assetid: 170f3e00-946c-c340-e12e-bee0078e62f3
ms.date: 06/08/2017
---


# AutoCaption.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[AutoCaption](autocaption-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[AutoCaption Object](autocaption-object-word.md)

