---
title: Selection.Style Property (Word)
keywords: vbawd10.chm158662664
f1_keywords:
- vbawd10.chm158662664
ms.prod: word
api_name:
- Word.Selection.Style
ms.assetid: d9295c79-97bd-3866-8321-45b708154716
ms.date: 06/08/2017
---


# Selection.Style Property (Word)

Returns or sets the style for the specified object. To set this property, specify the local name of the style, an integer, a  **WdBuiltinStyle** constant, or an object that represents the style. For a list of valid constants, consult the Microsoft Visual Basic Object Browser. Read/write **Variant** .


## Syntax

 _expression_ . **Style**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

When you return the style for a selection that includes more than one style, only the first character or paragraph style is returned.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

