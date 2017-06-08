---
title: OMathLimLow.Creator Property (Word)
keywords: vbawd10.chm91619429
f1_keywords:
- vbawd10.chm91619429
ms.prod: word
api_name:
- Word.OMathLimLow.Creator
ms.assetid: 634a993a-1d7c-c99d-112a-3e347331f71f
ms.date: 06/08/2017
---


# OMathLimLow.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[OMathLimLow](omathlimlow-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[OMathLimLow Object](omathlimlow-object-word.md)

