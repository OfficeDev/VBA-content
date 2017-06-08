---
title: Selection.DetectLanguage Method (Word)
keywords: vbawd10.chm158663191
f1_keywords:
- vbawd10.chm158663191
ms.prod: word
api_name:
- Word.Selection.DetectLanguage
ms.assetid: cfbc0d54-bb00-2bd0-ad9a-e646fdcbfe46
ms.date: 06/08/2017
---


# Selection.DetectLanguage Method (Word)

Analyzes the specified text to determine the language that it is written in.


## Syntax

 _expression_ . **DetectLanguage**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The results of the  **DetectLanguage** method are stored in the **LanguageID** property on a character-by-character basis. To read the **[LanguageID](language-id-property-word.md)** property, you must first specify a selection or range of text.

If a selection contains a partial sentence, the selection is extended to the end of the sentence.

If the  **DetectLanguage** method has already been applied to the specified text, the **LanguageDetected** property is set to **True** . To reevaulate the language of the specified text, you must first set the **[LanguageDetected](document-languagedetected-property-word.md)** property to **False** .


## See also


#### Concepts


[Selection Object](selection-object-word.md)

