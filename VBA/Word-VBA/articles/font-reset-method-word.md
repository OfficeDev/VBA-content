---
title: Font.Reset Method (Word)
keywords: vbawd10.chm156368998
f1_keywords:
- vbawd10.chm156368998
ms.prod: word
api_name:
- Word.Font.Reset
ms.assetid: 4e06c982-7b2b-03b2-b5c7-46370cb60044
ms.date: 06/08/2017
---


# Font.Reset Method (Word)

Removes manual character formatting (formatting not applied using a style). For example, if you manually format a word as bold and the underlying style is plain text (not bold), the  **Reset** method removes the bold format.


## Syntax

 _expression_ . **Reset**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Example

This example removes manual formatting from the selection.


```
Selection.Font.Reset
```


## See also


#### Concepts


[Font Object](font-object-word.md)

