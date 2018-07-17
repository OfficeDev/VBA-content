---
title: Options.ShowReadabilityStatistics Property (Word)
keywords: vbawd10.chm162988311
f1_keywords:
- vbawd10.chm162988311
ms.prod: word
api_name:
- Word.Options.ShowReadabilityStatistics
ms.assetid: 317a6175-75ea-f2eb-33ca-7eefd904e4c4
ms.date: 06/08/2017
---


# Options.ShowReadabilityStatistics Property (Word)

 **True** if Microsoft Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowReadabilityStatistics**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Word to show readability statistics, and then it checks the spelling and grammar in the active document.


```vb
Options.ShowReadabilityStatistics = True 
ActiveDocument.CheckGrammar
```

This example returns the current status of the Show readability statistics option on the Spelling &; Grammar tab in the Options dialog box (Tools menu).




```
temp = Options.ShowReadabilityStatistics
```


## See also


#### Concepts


[Options Object](options-object-word.md)

