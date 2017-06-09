---
title: OMathScrSup.Creator Property (Word)
keywords: vbawd10.chm119930981
f1_keywords:
- vbawd10.chm119930981
ms.prod: word
api_name:
- Word.OMathScrSup.Creator
ms.assetid: 6f7f3379-de41-2df9-17f8-252906726e58
ms.date: 06/08/2017
---


# OMathScrSup.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[OMathScrSup](omathscrsup-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[OMathScrSup Object](omathscrsup-object-word.md)

