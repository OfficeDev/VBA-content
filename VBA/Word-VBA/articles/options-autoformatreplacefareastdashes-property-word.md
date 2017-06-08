---
title: Options.AutoFormatReplaceFarEastDashes Property (Word)
keywords: vbawd10.chm162988327
f1_keywords:
- vbawd10.chm162988327
ms.prod: word
api_name:
- Word.Options.AutoFormatReplaceFarEastDashes
ms.assetid: 33b8c0c1-5249-05e6-d2a1-3565584207e5
ms.date: 06/08/2017
---


# Options.AutoFormatReplaceFarEastDashes Property (Word)

 **True** if long vowel sound and dash use is corrected when Microsoft Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplaceFarEastDashes**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically correct the use of long vowel sounds and dashes, and then it formats the current selection.


```vb
Options.AutoFormatReplaceFarEastDashes = True 
Selection.Range.AutoFormat
```


## See also


#### Concepts


[Options Object](options-object-word.md)

