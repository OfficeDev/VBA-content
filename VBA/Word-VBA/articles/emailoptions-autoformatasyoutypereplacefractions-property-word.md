---
title: EmailOptions.AutoFormatAsYouTypeReplaceFractions Property (Word)
keywords: vbawd10.chm165347595
f1_keywords:
- vbawd10.chm165347595
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplaceFractions
ms.assetid: 41a0273c-11c5-2053-fd7b-aaed13e1d9a1
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplaceFractions Property (Word)

 **True** if typed fractions are replaced with fractions from the current character set as you type; for example, "1/2" is replaced with "½." Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceFractions**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example turns off the automatic replacement of typed fractions.


```vb
Options.AutoFormatAsYouTypeReplaceFractions = False
```

This example returns the status of the  **Fractions (1/2) with fraction character (½)** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeReplaceFractions
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

