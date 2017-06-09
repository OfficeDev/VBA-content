---
title: Options.IgnoreUppercase Property (Word)
keywords: vbawd10.chm162988312
f1_keywords:
- vbawd10.chm162988312
ms.prod: word
api_name:
- Word.Options.IgnoreUppercase
ms.assetid: 4eff2832-3c66-0274-5403-d2fd8d31d04d
ms.date: 06/08/2017
---


# Options.IgnoreUppercase Property (Word)

 **True** if words in all uppercase letters are ignored while checking spelling. Read/write **Boolean** .


## Syntax

 _expression_ . **IgnoreUppercase**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Word to ignore words in all uppercase letters, and then it checks spelling in the active document.


```vb
Options.IgnoreUppercase = True 
ActiveDocument.CheckSpelling
```

This example returns the current status of the Ignore words in UPPERCASE option on the Spelling &; Grammar tab in the Options dialog box.




```vb
Dim blnTemp As Boolean 
 
blnTemp = Options.IgnoreUppercase
```


## See also


#### Concepts


[Options Object](options-object-word.md)

