---
title: Selection.ClearParagraphStyle Method (Word)
keywords: vbawd10.chm158663686
f1_keywords:
- vbawd10.chm158663686
ms.prod: word
api_name:
- Word.Selection.ClearParagraphStyle
ms.assetid: cfbafeac-99e1-5fae-a9a0-8cf8836add94
ms.date: 06/08/2017
---


# Selection.ClearParagraphStyle Method (Word)

Removes paragraph formatting that has been applied through paragraph styles from the selected text.


## Syntax

 _expression_ . **ClearParagraphStyle**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Remarks

This method does not remove paragraph formatting that a user has applied manually. To remove manually applied paragraph formatting, use the  **[ClearParagraphDirectFormatting](selection-clearparagraphdirectformatting-method-word.md)** method. To remove all paragraph formatting, both style and manual formatting, use the **[ClearParagraphAllFormatting](selection-clearparagraphallformatting-method-word.md)** method.


 **Note**  To remove character formatting, see the  **[ClearCharacterAllFormatting](selection-clearcharacterallformatting-method-word.md)** , **[ClearCharacterDirectFormatting](selection-clearcharacterdirectformatting-method-word.md)** , or **[ClearCharacterStyle](selection-clearcharacterstyle-method-word.md)** method.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

