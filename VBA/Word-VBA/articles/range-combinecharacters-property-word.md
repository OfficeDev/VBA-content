---
title: Range.CombineCharacters Property (Word)
keywords: vbawd10.chm157155595
f1_keywords:
- vbawd10.chm157155595
ms.prod: word
api_name:
- Word.Range.CombineCharacters
ms.assetid: 4852ebb7-b6cc-0bed-d1db-8a2efe14fc17
ms.date: 06/08/2017
---


# Range.CombineCharacters Property (Word)

 **True** if the specified range contains combined characters. Read/write **Boolean** .


## Syntax

 _expression_ . **CombineCharacters**

 _expression_ An expression that returns a **[Range](range-object-word.md)** object.


## Example

This example combines the characters in the selected range.


```vb
Selection.Range.CombineCharacters = True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

