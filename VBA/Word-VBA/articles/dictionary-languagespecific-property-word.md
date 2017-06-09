---
title: Dictionary.LanguageSpecific Property (Word)
keywords: vbawd10.chm162332677
f1_keywords:
- vbawd10.chm162332677
ms.prod: word
api_name:
- Word.Dictionary.LanguageSpecific
ms.assetid: 479eefb9-bd50-298b-635d-945ee7848600
ms.date: 06/08/2017
---


# Dictionary.LanguageSpecific Property (Word)

 **True** if the custom dictionary is to be used only with text formatted for a specific language. Read/write **Boolean** .


## Syntax

 _expression_ . **LanguageSpecific**

 _expression_ A variable that represents a **[Dictionary](dictionary-object-word.md)** object.


## Example

This example checks to see whether any custom dictionaries are language specific. If any of them are, the example removes them from the list of active custom dictionaries.


```vb
Dim dicLoop As Dictionary 
 
For each dicLoop in CustomDictionaries 
 If dicLoop.LanguageSpecific = True Then dicLoop.Delete 
Next dicLoop
```

This example adds a custom dictionary that will check only text that's formatted as German.




```vb
Dim dicNew As Dictionary 
 
Set dicNew = CustomDictionaries.Add("German1.dic") 
dicNew.LanguageSpecific = True 
dicNew.LanguageID = wdGerman
```


## See also


#### Concepts


[Dictionary Object](dictionary-object-word.md)

