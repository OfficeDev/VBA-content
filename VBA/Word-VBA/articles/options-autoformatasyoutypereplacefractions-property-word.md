---
title: Options.AutoFormatAsYouTypeReplaceFractions Property (Word)
keywords: vbawd10.chm162988299
f1_keywords:
- vbawd10.chm162988299
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeReplaceFractions
ms.assetid: fe741c60-b4dc-45ff-53d5-769b09a6b79b
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeReplaceFractions Property (Word)

 **True** if typed fractions are replaced with fractions from the current character set as you type. For example, "1/2" is replaced with "½." Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceFractions**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example turns off the automatic replacement of typed fractions.


```vb
Options.AutoFormatAsYouTypeReplaceFractions = False
```

This example returns the status of the Fractions (1/2) with fraction character (½) option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeReplaceFractions
```


## See also


#### Concepts


[Options Object](options-object-word.md)

