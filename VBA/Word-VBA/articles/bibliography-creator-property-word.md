---
title: Bibliography.Creator Property (Word)
keywords: vbawd10.chm101646437
f1_keywords:
- vbawd10.chm101646437
ms.prod: word
api_name:
- Word.Bibliography.Creator
ms.assetid: d836d894-c8f4-d6e6-c1de-e7c0f29aa918
ms.date: 06/08/2017
---


# Bibliography.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[Bibliography](bibliography-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[Bibliography Object](bibliography-object-word.md)

