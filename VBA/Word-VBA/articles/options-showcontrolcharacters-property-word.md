---
title: Options.ShowControlCharacters Property (Word)
keywords: vbawd10.chm162988438
f1_keywords:
- vbawd10.chm162988438
ms.prod: word
api_name:
- Word.Options.ShowControlCharacters
ms.assetid: 9fed5e7a-79b9-0517-e985-7d53a642220c
ms.date: 06/08/2017
---


# Options.ShowControlCharacters Property (Word)

 **True** if bidirectional control characters are visible in the current document. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowControlCharacters**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example hides bidirectional control characters in the current document.


```vb
Options.ShowControlCharacters = False
```


## See also


#### Concepts


[Options Object](options-object-word.md)

