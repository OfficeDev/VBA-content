---
title: Options.ConvertHighAnsiToFarEast Property (Word)
keywords: vbawd10.chm162988360
f1_keywords:
- vbawd10.chm162988360
ms.prod: word
api_name:
- Word.Options.ConvertHighAnsiToFarEast
ms.assetid: d973f327-1887-cca8-344a-80ce3f9e740a
ms.date: 06/08/2017
---


# Options.ConvertHighAnsiToFarEast Property (Word)

 **True** if Microsoft Word converts text that is associated with an East Asian font to the appropriate font when it opens a document. Read/write **Boolean** .


## Syntax

 _expression_ . **ConvertHighAnsiToFarEast**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to convert text that is associated with an East Asian font to the appropriate font when it opens a document.


```vb
Options.ConvertHighAnsiToFarEast = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)

