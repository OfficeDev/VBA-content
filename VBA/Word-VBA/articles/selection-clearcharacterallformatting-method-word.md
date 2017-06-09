---
title: Selection.ClearCharacterAllFormatting Method (Word)
keywords: vbawd10.chm158663687
f1_keywords:
- vbawd10.chm158663687
ms.prod: word
api_name:
- Word.Selection.ClearCharacterAllFormatting
ms.assetid: 1d0dfb43-4855-1534-5ec2-475232a6a457
ms.date: 06/08/2017
---


# Selection.ClearCharacterAllFormatting Method (Word)

Removes all character formatting (formatting applied either through character styles or manually applied formatting) from the selected text.


## Syntax

 _expression_ . **ClearCharacterAllFormatting**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Remarks

This method removes all character formatting. If you need to removed formatting applied through character styles, use the  **[ClearCharacterStyle](selection-clearcharacterstyle-method-word.md)** method. To remove character formatting that the user has manually applied using Microsoft Word character formatting features, use the **[ClearCharacterDirectFormatting](selection-clearcharacterdirectformatting-method-word.md)** method.


 **Note**  To remove paragraph formatting, see the  **[ClearParagraphAllFormatting](selection-clearparagraphallformatting-method-word.md)** , **[ClearParagraphDirectFormatting](selection-clearparagraphdirectformatting-method-word.md)** , or **[ClearParagraphStyle](selection-clearparagraphstyle-method-word.md)** method.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

