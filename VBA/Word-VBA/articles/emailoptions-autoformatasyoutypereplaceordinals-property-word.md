---
title: EmailOptions.AutoFormatAsYouTypeReplaceOrdinals Property (Word)
keywords: vbawd10.chm165347594
f1_keywords:
- vbawd10.chm165347594
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplaceOrdinals
ms.assetid: c3f60ea8-1949-6247-98d1-d8d386507341
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplaceOrdinals Property (Word)

 **True** if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript as you type; for example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceOrdinals**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example turns on the automatic replacement of ordinals with superscript letters.


```vb
Options.AutoFormatAsYouTypeReplaceOrdinals = True
```

This example returns the status of the  **Ordinals (1st) with superscript** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeReplaceOrdinals
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

