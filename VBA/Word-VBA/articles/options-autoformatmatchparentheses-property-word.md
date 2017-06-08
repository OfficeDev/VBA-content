---
title: Options.AutoFormatMatchParentheses Property (Word)
keywords: vbawd10.chm162988326
f1_keywords:
- vbawd10.chm162988326
ms.prod: word
api_name:
- Word.Options.AutoFormatMatchParentheses
ms.assetid: edc8901c-6eb2-bb89-5054-3ed4888d2199
ms.date: 06/08/2017
---


# Options.AutoFormatMatchParentheses Property (Word)

 **True** if improperly paired parentheses are corrected when Microsoft Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatMatchParentheses**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically correct pairs of parentheses, and then it formats the current selection.


```vb
Options.AutoFormatMatchParentheses = True 
Selection.Range.AutoFormat
```


## See also


#### Concepts


[Options Object](options-object-word.md)

