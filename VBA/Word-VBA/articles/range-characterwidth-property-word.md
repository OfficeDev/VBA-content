---
title: Range.CharacterWidth Property (Word)
keywords: vbawd10.chm157155654
f1_keywords:
- vbawd10.chm157155654
ms.prod: word
api_name:
- Word.Range.CharacterWidth
ms.assetid: 83eadb2b-5c79-d246-d1f1-fd6a9e1f4bd8
ms.date: 06/08/2017
---


# Range.CharacterWidth Property (Word)

Returns or sets the character width of the specified range. Read/write  **WdCharacterWidth** .


## Syntax

 _expression_ . **CharacterWidth**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example converts the current selection to half-width characters.


```
Selection.Range.CharacterWidth = wdWidthHalfWidth
```


## See also


#### Concepts


[Range Object](range-object-word.md)

