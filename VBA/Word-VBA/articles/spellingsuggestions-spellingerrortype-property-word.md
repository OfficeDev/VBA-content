---
title: SpellingSuggestions.SpellingErrorType Property (Word)
keywords: vbawd10.chm162136066
f1_keywords:
- vbawd10.chm162136066
ms.prod: word
api_name:
- Word.SpellingSuggestions.SpellingErrorType
ms.assetid: 0d5c71a3-77eb-b36c-76b8-c6fd49bb6394
ms.date: 06/08/2017
---


# SpellingSuggestions.SpellingErrorType Property (Word)

Returns the spelling error type. Read-only  **WdSpellingErrorType** .


## Syntax

 _expression_ . **SpellingErrorType**

 _expression_ Required. A variable that represents a **[SpellingSuggestions](spellingsuggestions-object-word.md)** collection.


## Remarks

Use the  **GetSpellingSuggestions** method to return a collection of words suggested as spelling replacements. If a word is misspelled, the **CheckSpelling** method returns **True** .


## Example

If the first word in the active document isn't in the dictionary, this example displays "Unknown word" in the status bar.


```vb
Set suggs = ActiveDocument.Content.GetSpellingSuggestions 
If suggs.SpellingErrorType = wdSpellingNotInDictionary Then 
 StatusBar = "Unknown word" 
End If
```


## See also


#### Concepts


[SpellingSuggestions Collection Object](spellingsuggestions-object-word.md)

