---
title: Selection.ClearCharacterDirectFormatting Method (Word)
keywords: vbawd10.chm158663689
f1_keywords:
- vbawd10.chm158663689
ms.prod: word
api_name:
- Word.Selection.ClearCharacterDirectFormatting
ms.assetid: d2138876-c832-2407-a53e-5bd4af2421b7
ms.date: 06/08/2017
---


# Selection.ClearCharacterDirectFormatting Method (Word)

Removes character formatting (formatting that has been applied manually using the buttons on the ribbon or through the dialog boxes) from the selected text.


## Syntax

 _expression_ . **ClearCharacterDirectFormatting**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Remarks

This method does not remove character formatting that has been applied by using a character style. To remove character formatting that the user has applied by using character styles, use the  **[ClearCharacterStyle](selection-clearcharacterstyle-method-word.md)** method. To remove all character formatting, regardless of whether the user has applied it by using character styles or manually through the formatting features in Microsoft Word, use the **[ClearCharacterAllFormatting](selection-clearcharacterallformatting-method-word.md)** method.


 **Note**  For more information about how to remove paragraph formatting, see the  **[ClearParagraphAllFormatting](selection-clearparagraphallformatting-method-word.md)** , **[ClearParagraphDirectFormatting](selection-clearparagraphdirectformatting-method-word.md)** , or **[ClearParagraphStyle](selection-clearparagraphstyle-method-word.md)** method.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

