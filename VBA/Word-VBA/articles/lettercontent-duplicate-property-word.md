---
title: LetterContent.Duplicate Property (Word)
keywords: vbawd10.chm161546250
f1_keywords:
- vbawd10.chm161546250
ms.prod: word
api_name:
- Word.LetterContent.Duplicate
ms.assetid: 925ba556-4a7e-36da-2fbb-a32684f23fa6
ms.date: 06/08/2017
---


# LetterContent.Duplicate Property (Word)

Returns a read-only  **LetterContent** object that represents the contents of a letter created by the Letter Wizard.


## Syntax

 _expression_ . **Duplicate**

 _expression_ Required. A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Remarks

You can use the  **Duplicate** property to pick up the settings of all the properties of a duplicated **LetterContent** object. You can assign the object returned by the **Duplicate** property to another **LetterContent** object to apply those settings all at once. Before assigning the duplicate object to another object, you can change any of the properties of the duplicate object without affecting the original.


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

