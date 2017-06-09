---
title: Selection.LanguageID Property (Word)
keywords: vbawd10.chm158662809
f1_keywords:
- vbawd10.chm158662809
ms.prod: word
api_name:
- Word.Selection.LanguageID
ms.assetid: d92be532-99db-8b46-3e64-8a3fca65004e
ms.date: 06/08/2017
---


# Selection.LanguageID Property (Word)

Returns or sets the language for the specified object. Read/write .


## Syntax

 _expression_ . **LanguageID**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For a custom dictionary, you must first set the  **LanguageSpecific** property to **True** before specifying the **LanguageID** property. Custom dictionaries that are language-specific check only text that is formatted for that language.

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## See also


#### Concepts


[Selection Object](selection-object-word.md)

