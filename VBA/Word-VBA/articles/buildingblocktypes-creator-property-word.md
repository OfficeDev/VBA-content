---
title: BuildingBlockTypes.Creator Property (Word)
keywords: vbawd10.chm189793257
f1_keywords:
- vbawd10.chm189793257
ms.prod: word
api_name:
- Word.BuildingBlockTypes.Creator
ms.assetid: e8b9c1b2-542e-9f9b-8100-51c82d886eea
ms.date: 06/08/2017
---


# BuildingBlockTypes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[BuildingBlockTypes](buildingblocktypes-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[BuildingBlockTypes Collection](buildingblocktypes-object-word.md)

