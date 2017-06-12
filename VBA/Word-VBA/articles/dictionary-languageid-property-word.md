---
title: Dictionary.LanguageID Property (Word)
keywords: vbawd10.chm162332674
f1_keywords:
- vbawd10.chm162332674
ms.prod: word
api_name:
- Word.Dictionary.LanguageID
ms.assetid: 598efc88-f26d-49b2-6451-e2cbedd20ff7
ms.date: 06/08/2017
---


# Dictionary.LanguageID Property (Word)

Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified object. Read/write.


## Syntax

 _expression_**LanguageID**

 _expression_ Required. An expression that returns a **[Dictionary](dictionary-object-word.md)** object.


## Remarks

For a custom dictionary, you must first set the  **[LanguageSpecific](dictionary-languagespecific-property-word.md)** property to **True** before specifying the **LanguageID** property. Custom dictionaries that are language-specific check only text that is formatted for that language.

Some  **WdLanguageID** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## See also


#### Concepts


[Dictionary Object](dictionary-object-word.md)

