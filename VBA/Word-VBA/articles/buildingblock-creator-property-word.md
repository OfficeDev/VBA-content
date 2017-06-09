---
title: BuildingBlock.Creator Property (Word)
keywords: vbawd10.chm203621353
f1_keywords:
- vbawd10.chm203621353
ms.prod: word
api_name:
- Word.BuildingBlock.Creator
ms.assetid: 97f89a5d-3a4a-63a8-12bc-086a864d80c8
ms.date: 06/08/2017
---


# BuildingBlock.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[BuildingBlock](buildingblock-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[BuildingBlock Object](buildingblock-object-word.md)

