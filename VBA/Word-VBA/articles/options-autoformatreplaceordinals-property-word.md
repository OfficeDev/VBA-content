---
title: Options.AutoFormatReplaceOrdinals Property (Word)
keywords: vbawd10.chm162988288
f1_keywords:
- vbawd10.chm162988288
ms.prod: word
api_name:
- Word.Options.AutoFormatReplaceOrdinals
ms.assetid: 7dd6d253-53e5-5fec-aafa-181899afe02b
ms.date: 06/08/2017
---


# Options.AutoFormatReplaceOrdinals Property (Word)

 **True** if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript when Word formats a document or range automatically. For example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplaceOrdinals**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example turns on the automatic replacement of ordinals with superscript, and then it formats the current selection automatically.


```vb
Options.AutoFormatReplaceOrdinals = True 
Selection.Range.AutoFormat
```

This example returns the status of the Ordinals (1st) with superscript option on the  **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceOrdinals
```


## See also


#### Concepts


[Options Object](options-object-word.md)

