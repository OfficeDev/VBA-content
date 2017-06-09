---
title: Replacement.Style Property (Word)
keywords: vbawd10.chm162594828
f1_keywords:
- vbawd10.chm162594828
ms.prod: word
api_name:
- Word.Replacement.Style
ms.assetid: 4cf38f58-e50b-d39c-18f7-4fb35c8c0575
ms.date: 06/08/2017
---


# Replacement.Style Property (Word)

Returns or sets the style for the specified object. To set this property, specify the local name of the style, an integer, a  **[WdBuiltinStyle](wdbuiltinstyle-enumeration-word.md)** constant, or an object that represents the style. Read/write **Variant** .


## Syntax

 _expression_ . **Style**

 _expression_ Required. A variable that represents a **[Replacement](replacement-object-word.md)** object.


## Remarks

When you return the style for a range that includes more than one style, only the first character or paragraph style is returned.


## See also


#### Concepts


[Replacement Object](replacement-object-word.md)

