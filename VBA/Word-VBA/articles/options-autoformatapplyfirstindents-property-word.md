---
title: Options.AutoFormatApplyFirstIndents Property (Word)
keywords: vbawd10.chm162988323
f1_keywords:
- vbawd10.chm162988323
ms.prod: word
api_name:
- Word.Options.AutoFormatApplyFirstIndents
ms.assetid: c55fa4eb-9ef4-9061-b2be-cbe2da8ce3bf
ms.date: 06/08/2017
---


# Options.AutoFormatApplyFirstIndents Property (Word)

 **True** if Microsoft Word replaces a space entered at the beginning of a paragraph with a first-line indent when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatApplyFirstIndents**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to replace a space entered at the beginning of a paragraph with a first-line indent and automatically formats the selected range.


```vb
Options.AutoFormatApplyFirstIndents = True 
Selection.Range.AutoFormat
```


## See also


#### Concepts


[Options Object](options-object-word.md)

