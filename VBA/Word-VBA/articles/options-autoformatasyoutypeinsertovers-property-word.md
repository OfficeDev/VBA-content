---
title: Options.AutoFormatAsYouTypeInsertOvers Property (Word)
keywords: vbawd10.chm162988337
f1_keywords:
- vbawd10.chm162988337
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeInsertOvers
ms.assetid: e79cd972-85c3-aa9a-abab-a92ceb171213
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeInsertOvers Property (Word)

 **True** for Microsoft Word to automatically insert "以上" when the user enters "記" or "案". Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeInsertOvers**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically insert "以上" when the user enters "記" or "案".


```vb
Options.AutoFormatAsYouTypeInsertOvers = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)

