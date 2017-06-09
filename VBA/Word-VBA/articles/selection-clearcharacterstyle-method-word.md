---
title: Selection.ClearCharacterStyle Method (Word)
keywords: vbawd10.chm158663688
f1_keywords:
- vbawd10.chm158663688
ms.prod: word
api_name:
- Word.Selection.ClearCharacterStyle
ms.assetid: ff9795f9-ea74-fa03-5d87-9c56152d179d
ms.date: 06/08/2017
---


# Selection.ClearCharacterStyle Method (Word)

Removes character formatting that has been applied through character styles from the selected text.


## Syntax

 _expression_ . **ClearCharacterStyle**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Remarks

This method does not remove character formatting that a user has applied manually. To remove manually applied character formatting, use the  **[ClearCharacterDirectFormatting](selection-clearcharacterdirectformatting-method-word.md)** method. To remove all character formatting, both style and manual formatting, use the **[ClearCharacterAllFormatting](selection-clearcharacterallformatting-method-word.md)** method.


 **Note**  To remove paragraph formatting, see the  **[ClearParagraphAllFormatting](selection-clearparagraphallformatting-method-word.md)** , **[ClearParagraphDirectFormatting](selection-clearparagraphdirectformatting-method-word.md)** , or **[ClearParagraphStyle](selection-clearparagraphstyle-method-word.md)** method.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

