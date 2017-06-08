---
title: OMathBox.OpEmu Property (Word)
keywords: vbawd10.chm134086760
f1_keywords:
- vbawd10.chm134086760
ms.prod: word
api_name:
- Word.OMathBox.OpEmu
ms.assetid: 27e17879-b26b-cdc0-87fd-e947942ac97b
ms.date: 06/08/2017
---


# OMathBox.OpEmu Property (Word)

Returns or sets a  **Boolean** that represents that the box and its contents behave as a single operator and inherit the properties of an operator. Read/write.


## Syntax

 _expression_ . **OpEmu**

 _expression_ An expression that returns an **[OMathBox](omathbox-object-word.md)** object.


## Remarks

When the OpEmu property is  **True** , the character can serve as a point for a line break and can be aligned to other operators. Operator emulators are often used when one or more glyphs combine to form an operator, such as ==.


## See also


#### Concepts


[OMathBox Object](omathbox-object-word.md)

