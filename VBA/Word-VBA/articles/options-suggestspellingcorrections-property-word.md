---
title: Options.SuggestSpellingCorrections Property (Word)
keywords: vbawd10.chm162988315
f1_keywords:
- vbawd10.chm162988315
ms.prod: word
api_name:
- Word.Options.SuggestSpellingCorrections
ms.assetid: 2b4e821a-f44b-9166-5cf9-ff607164a99c
ms.date: 06/08/2017
---


# Options.SuggestSpellingCorrections Property (Word)

 **True** if Microsoft Word always suggests alternative spellings for each misspelled word when checking spelling. Read/write **Boolean** .


## Syntax

 _expression_ . **SuggestSpellingCorrections**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Word to always suggest alternative spellings for misspelled words, and then it checks spelling in the active document.


```vb
Options.SuggestSpellingCorrections = True 
ActiveDocument.CheckSpelling
```

This example returns the current status of the Always suggest corrections option on the  **Spelling &; Grammar** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.SuggestSpellingCorrections
```


## See also


#### Concepts


[Options Object](options-object-word.md)

