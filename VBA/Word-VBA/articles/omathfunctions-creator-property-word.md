---
title: OMathFunctions.Creator Property (Word)
keywords: vbawd10.chm44302437
f1_keywords:
- vbawd10.chm44302437
ms.prod: word
api_name:
- Word.OMathFunctions.Creator
ms.assetid: c95e5aa0-654c-25b6-0baf-43592ba159c3
ms.date: 06/08/2017
---


# OMathFunctions.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[OMathFunctions](omathfunctions-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[OMathFunctions Collection](omathfunctions-object-word.md)

