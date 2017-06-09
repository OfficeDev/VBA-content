---
title: Options.IgnoreMixedDigits Property (Word)
keywords: vbawd10.chm162988313
f1_keywords:
- vbawd10.chm162988313
ms.prod: word
api_name:
- Word.Options.IgnoreMixedDigits
ms.assetid: 3603afd8-a922-dec6-2239-6ae1d330995e
ms.date: 06/08/2017
---


# Options.IgnoreMixedDigits Property (Word)

 **True** if words that contain numbers are ignored while checking spelling. Read/write **Boolean** .


## Syntax

 _expression_ . **IgnoreMixedDigits**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore words that contain numbers, and then it checks spelling in the active document.


```vb
Options.IgnoreMixedDigits = True 
ActiveDocument.CheckSpelling
```

This example returns the current status of the Ignore words with numbers option on the Spelling &; Grammar tab in the Options dialog box.




```vb
Dim blnTemp As Boolean 
 
blnTemp = Options.IgnoreMixedDigits
```


## See also


#### Concepts


[Options Object](options-object-word.md)

