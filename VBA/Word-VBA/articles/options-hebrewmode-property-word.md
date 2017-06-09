---
title: Options.HebrewMode Property (Word)
keywords: vbawd10.chm162988443
f1_keywords:
- vbawd10.chm162988443
ms.prod: word
api_name:
- Word.Options.HebrewMode
ms.assetid: 8a98159e-099d-299c-c955-2190d683d450
ms.date: 06/08/2017
---


# Options.HebrewMode Property (Word)

Returns or sets the mode for the Hebrew spelling checker. Read/write  **WdHebSpellStart** .


## Syntax

 _expression_ . **HebrewMode**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the spelling checker to check spelling based on the conventional script required by the Hebrew Language Academy for writing text with diacritics.


```
Options.HebrewMode = wdFullScript
```


## See also


#### Concepts


[Options Object](options-object-word.md)

