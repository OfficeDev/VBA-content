---
title: Options.AutoFormatAsYouTypeReplaceOrdinals Property (Word)
keywords: vbawd10.chm162988298
f1_keywords:
- vbawd10.chm162988298
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeReplaceOrdinals
ms.assetid: eebf3119-8743-834f-7425-5adc60a1a7ef
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeReplaceOrdinals Property (Word)

 **True** if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript as you type. For example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceOrdinals**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example turns on the automatic replacement of ordinals with superscript letters.


```vb
Options.AutoFormatAsYouTypeReplaceOrdinals = True
```

This example returns the status of the Ordinals (1st) with superscript option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeReplaceOrdinals
```


## See also


#### Concepts


[Options Object](options-object-word.md)

