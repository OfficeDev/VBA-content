---
title: Selection.ClearParagraphDirectFormatting Method (Word)
keywords: vbawd10.chm158663696
f1_keywords:
- vbawd10.chm158663696
ms.prod: word
api_name:
- Word.Selection.ClearParagraphDirectFormatting
ms.assetid: 66df2319-f02e-7cd9-4cef-fda6468dcd67
ms.date: 06/08/2017
---


# Selection.ClearParagraphDirectFormatting Method (Word)

Removes paragraph formatting that has been applied manually (using the buttons on the ribbon or through the dialog boxes) from the selected text.


## Syntax

 _expression_ . **ClearParagraphDirectFormatting**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Remarks

This method does not remove paragraph formatting that has been applied by using a paragraph style. To remove paragraph formatting that the user has applied by using paragraph styles, use the  **[ClearParagraphStyle](selection-clearparagraphstyle-method-word.md)** method. To remove all paragraph formatting, both style and manual formatting, use the **[ClearParagraphAllFormatting](selection-clearparagraphallformatting-method-word.md)** method.


 **Note**  For more information about how to remove character formatting, see the  **[ClearCharacterAllFormatting](selection-clearcharacterallformatting-method-word.md)** , **[ClearCharacterDirectFormatting](selection-clearcharacterdirectformatting-method-word.md)** , or **[ClearCharacterStyle](selection-clearcharacterstyle-method-word.md)** method.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

